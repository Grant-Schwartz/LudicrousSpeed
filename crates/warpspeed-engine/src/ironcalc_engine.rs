use std::collections::{HashMap, HashSet, VecDeque};
use std::fs;
use std::sync::Mutex;
use std::time::Instant;

use ironcalc::base::{cell::CellValue, expressions::utils::number_to_column, Model};
use ironcalc::import::load_from_xlsx;
use serde_json::json;
use sha2::{Digest, Sha256};

use crate::data_tables::{
    evaluate_data_tables, evaluate_iterative_formula_cells, summarize_data_tables, DataTableRegion,
};
use crate::engine::CalcEngine;
use crate::graph::{build_dependency_graph, CellId, DependencyGraph, DirtySummary};
use crate::model::{
    AnalysisSummary, BenchmarkSummary, CalcMode, CalcPlan, CalcResult, CalculationStrategy,
    ChangedCell, DataTableBenchmarkSummary, DataTableEvaluationStatus, EngineError,
    ExcelWritebackPlan, FallbackDetail, FallbackReason, FormulaCoverage, FormulaValueKind,
    FormulaWritebackCell, WorkbookSnapshot, WritebackIssueSummary, WritebackMode,
};
use crate::xlsx_sanitize::{
    remove_sanitized_workbook, sanitize_data_table_formulas, ImportFallbacks,
};

const CACHE_CAPACITY: usize = 4;
const CACHE_LANGUAGE: &str = "en";

pub struct IronCalcEngine {
    cache: Mutex<EngineCache>,
}

impl Default for IronCalcEngine {
    fn default() -> Self {
        Self {
            cache: Mutex::new(EngineCache::new(CACHE_CAPACITY)),
        }
    }
}

impl std::fmt::Debug for IronCalcEngine {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("IronCalcEngine").finish_non_exhaustive()
    }
}

impl CalcEngine for IronCalcEngine {
    fn plan(&self, snapshot: &WorkbookSnapshot) -> Result<CalcPlan, EngineError> {
        Ok(CalcPlan {
            workbook_hash: String::new(),
            mode: snapshot.mode,
            fallback_reasons: Vec::new(),
        })
    }

    fn calculate(
        &self,
        snapshot: &WorkbookSnapshot,
        mut plan: CalcPlan,
    ) -> Result<CalcResult, EngineError> {
        let started = Instant::now();
        let cache_started = Instant::now();
        let cache_key = cache_key(snapshot);
        let mut cache = self
            .cache
            .lock()
            .map_err(|_| EngineError::Evaluation("engine cache lock is poisoned".to_string()))?;
        let cache_lookup_ms = cache_started.elapsed().as_millis();

        if let Some(key) = cache_key.as_deref() {
            if !snapshot.force_reload {
                if let Some(mut entry) = cache.take(key) {
                    if snapshot.changed_cells.is_empty()
                        && entry.is_result_cache_safe()
                        && entry.last_result.is_some()
                    {
                        let mut result = entry.last_result.clone().unwrap();
                        let mut data_tables =
                            summarize_data_tables(&entry.model, &entry.data_tables, &[], false);
                        if data_tables.data_table_count > 0 {
                            data_tables.status = DataTableEvaluationStatus::Reused;
                        }
                        result.plan.mode = snapshot.mode;
                        result.plan.workbook_hash = entry.workbook_hash.clone();
                        result.plan.fallback_reasons = merged_fallback_reasons(
                            &entry.graph,
                            &entry.import_fallbacks,
                            &data_tables,
                        );
                        result.writeback = build_writeback_plan(
                            snapshot,
                            &entry.model,
                            &entry.graph,
                            &entry.import_fallbacks,
                            &entry.data_tables,
                            &data_tables,
                            &entry.cached_formula_values,
                        );
                        result.benchmark = benchmark_summary(
                            snapshot,
                            started,
                            TimingBreakdown {
                                cache_lookup_ms,
                                model_cache_hit: true,
                                graph_cache_hit: true,
                                result_cache_hit: true,
                                strategy: CalculationStrategy::WarmNoopCacheHit,
                                cache_hit_rate: 1.0,
                                data_tables,
                                ..TimingBreakdown::default()
                            },
                        );
                        cache.insert(key.to_string(), entry);
                        return Ok(result);
                    }

                    let applied =
                        match apply_changed_cells(&mut entry.model, &snapshot.changed_cells) {
                            Ok(applied) => applied,
                            Err(err) if !snapshot.workbook_path.trim().is_empty() => {
                                return run_forced_reload(
                                    &mut cache,
                                    key,
                                    snapshot,
                                    plan,
                                    started,
                                    cache_lookup_ms,
                                    err,
                                );
                            }
                            Err(err) => return Err(err),
                        };

                    let mut graph_build_ms = 0;
                    let graph_cache_hit =
                        !applied.formula_changed && !entry.graph.has_cached_circular_allowances();
                    if applied.formula_changed || entry.graph.has_cached_circular_allowances() {
                        let graph_started = Instant::now();
                        entry.graph = build_dependency_graph(&entry.model);
                        entry.structure_hash = entry.graph.structure_hash().to_string();
                        graph_build_ms = graph_started.elapsed().as_millis();
                    }

                    let dirty = entry.graph.dirty_summary(&applied.changed_cells);
                    entry.cached_formula_values =
                        collect_circular_formula_values(&entry.model, &entry.graph, false);
                    let circular_value_cells = entry
                        .cached_formula_values
                        .keys()
                        .copied()
                        .collect::<HashSet<_>>();
                    entry
                        .graph
                        .allow_cached_circular_values(&circular_value_cells, &[]);
                    let eval_started = Instant::now();
                    entry.model.evaluate();
                    let ironcalc_ms = eval_started.elapsed().as_millis();
                    let data_tables = evaluate_data_tables(
                        &entry.model,
                        &entry.data_tables,
                        &applied.changed_cells,
                        snapshot.evaluate_data_tables,
                    );

                    let result = build_result(
                        snapshot,
                        &mut plan,
                        &entry.workbook_hash,
                        &entry.graph,
                        &entry.import_fallbacks,
                        &entry.data_tables,
                        &entry.model,
                        &entry.cached_formula_values,
                        started,
                        TimingBreakdown {
                            cache_lookup_ms,
                            graph_build_ms,
                            ironcalc_ms,
                            model_cache_hit: true,
                            graph_cache_hit,
                            strategy: CalculationStrategy::WarmFullWithDirtyPlan,
                            cache_hit_rate: cache_hit_rate(&entry.graph, dirty),
                            dirty,
                            data_tables,
                            ..TimingBreakdown::default()
                        },
                    );
                    entry.last_result = Some(result.clone());
                    cache.insert(key.to_string(), entry);
                    return Ok(result);
                }
            }
        }

        let key = cache_key.unwrap_or_else(|| snapshot.workbook_path.clone());
        let strategy = if snapshot.force_reload {
            CalculationStrategy::ForcedReload
        } else {
            CalculationStrategy::ColdFull
        };
        let entry = load_build_evaluate(snapshot)?;
        let result = build_result(
            snapshot,
            &mut plan,
            &entry.workbook_hash,
            &entry.graph,
            &entry.import_fallbacks,
            &entry.data_tables,
            &entry.model,
            &entry.cached_formula_values,
            started,
            TimingBreakdown {
                cache_lookup_ms,
                load_ms: entry.load_ms,
                graph_build_ms: entry.graph_build_ms,
                ironcalc_ms: entry.ironcalc_ms,
                strategy,
                data_tables: entry.data_table_summary.clone(),
                ..TimingBreakdown::default()
            },
        );
        let mut cached_entry = entry.into_cached();
        cached_entry.last_result = Some(result.clone());
        cache.insert(key, cached_entry);
        Ok(result)
    }
}

struct EngineCache {
    capacity: usize,
    entries: HashMap<String, CachedWorkbook>,
    lru: VecDeque<String>,
}

impl EngineCache {
    fn new(capacity: usize) -> Self {
        Self {
            capacity,
            entries: HashMap::new(),
            lru: VecDeque::new(),
        }
    }

    fn take(&mut self, key: &str) -> Option<CachedWorkbook> {
        let entry = self.entries.remove(key);
        if entry.is_some() {
            self.lru.retain(|existing| existing != key);
        }
        entry
    }

    fn insert(&mut self, key: String, entry: CachedWorkbook) {
        self.entries.insert(key.clone(), entry);
        self.lru.retain(|existing| existing != &key);
        self.lru.push_back(key);
        while self.entries.len() > self.capacity {
            let Some(oldest) = self.lru.pop_front() else {
                break;
            };
            self.entries.remove(&oldest);
        }
    }
}

struct CachedWorkbook {
    model: Model<'static>,
    graph: DependencyGraph,
    import_fallbacks: ImportFallbacks,
    data_tables: Vec<DataTableRegion>,
    cached_formula_values: HashMap<CellId, CachedFormulaValue>,
    structure_hash: String,
    workbook_hash: String,
    last_result: Option<CalcResult>,
}

impl CachedWorkbook {
    fn is_result_cache_safe(&self) -> bool {
        self.graph.is_result_cache_safe() && self.import_fallbacks.is_empty()
    }
}

struct LoadedWorkbook {
    model: Model<'static>,
    graph: DependencyGraph,
    import_fallbacks: ImportFallbacks,
    data_tables: Vec<DataTableRegion>,
    cached_formula_values: HashMap<CellId, CachedFormulaValue>,
    structure_hash: String,
    workbook_hash: String,
    load_ms: u128,
    graph_build_ms: u128,
    ironcalc_ms: u128,
    data_table_summary: DataTableBenchmarkSummary,
}

impl LoadedWorkbook {
    fn into_cached(self) -> CachedWorkbook {
        CachedWorkbook {
            model: self.model,
            graph: self.graph,
            import_fallbacks: self.import_fallbacks,
            data_tables: self.data_tables,
            cached_formula_values: self.cached_formula_values,
            structure_hash: self.structure_hash,
            workbook_hash: self.workbook_hash,
            last_result: None,
        }
    }
}

#[derive(Default)]
struct TimingBreakdown {
    cache_lookup_ms: u128,
    load_ms: u128,
    graph_build_ms: u128,
    ironcalc_ms: u128,
    model_cache_hit: bool,
    graph_cache_hit: bool,
    result_cache_hit: bool,
    strategy: CalculationStrategy,
    dirty: DirtySummary,
    cache_hit_rate: f64,
    data_tables: DataTableBenchmarkSummary,
}

impl Default for CalculationStrategy {
    fn default() -> Self {
        CalculationStrategy::ColdFull
    }
}

struct AppliedChanges {
    changed_cells: Vec<CellId>,
    formula_changed: bool,
}

#[derive(Clone, Debug)]
struct CachedFormulaValue {
    value_kind: FormulaValueKind,
    value: serde_json::Value,
}

fn cache_key(snapshot: &WorkbookSnapshot) -> Option<String> {
    snapshot
        .workbook_id
        .as_deref()
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(ToOwned::to_owned)
        .or_else(|| {
            let path = snapshot.workbook_path.trim();
            (!path.is_empty()).then(|| path.to_string())
        })
}

fn run_forced_reload(
    cache: &mut EngineCache,
    key: &str,
    snapshot: &WorkbookSnapshot,
    mut plan: CalcPlan,
    started: Instant,
    cache_lookup_ms: u128,
    original_error: EngineError,
) -> Result<CalcResult, EngineError> {
    if snapshot.workbook_path.trim().is_empty() {
        return Err(original_error);
    }

    let entry = load_build_evaluate(snapshot)?;
    let result = build_result(
        snapshot,
        &mut plan,
        &entry.workbook_hash,
        &entry.graph,
        &entry.import_fallbacks,
        &entry.data_tables,
        &entry.model,
        &entry.cached_formula_values,
        started,
        TimingBreakdown {
            cache_lookup_ms,
            load_ms: entry.load_ms,
            graph_build_ms: entry.graph_build_ms,
            ironcalc_ms: entry.ironcalc_ms,
            strategy: CalculationStrategy::ForcedReload,
            data_tables: entry.data_table_summary.clone(),
            ..TimingBreakdown::default()
        },
    );
    let mut cached_entry = entry.into_cached();
    cached_entry.last_result = Some(result.clone());
    cache.insert(key.to_string(), cached_entry);
    Ok(result)
}

fn load_build_evaluate(snapshot: &WorkbookSnapshot) -> Result<LoadedWorkbook, EngineError> {
    let workbook_path = snapshot.workbook_path.trim();
    if workbook_path.is_empty() {
        return Err(EngineError::WorkbookLoad(
            "cache miss requires workbook_path".to_string(),
        ));
    }

    let load_started = Instant::now();
    let bytes =
        fs::read(workbook_path).map_err(|err| EngineError::WorkbookLoad(err.to_string()))?;
    let workbook_hash = format!("{:x}", Sha256::digest(&bytes));
    let (mut model, import_fallbacks, data_tables) = load_model_with_import_fallbacks(
        workbook_path,
        &workbook_hash,
        &snapshot.locale,
        &snapshot.timezone,
    )?;
    let load_ms = load_started.elapsed().as_millis();

    let graph_started = Instant::now();
    let mut graph = build_dependency_graph(&model);
    let cached_formula_values = collect_circular_formula_values(&model, &graph, true);
    let cached_value_cells = cached_formula_values
        .keys()
        .copied()
        .collect::<HashSet<_>>();
    graph.allow_cached_circular_values(&cached_value_cells, &[]);
    let structure_hash = graph.structure_hash().to_string();
    let graph_build_ms = graph_started.elapsed().as_millis();

    let eval_started = Instant::now();
    model.evaluate();
    let ironcalc_ms = eval_started.elapsed().as_millis();
    let data_table_summary =
        evaluate_data_tables(&model, &data_tables, &[], snapshot.evaluate_data_tables);

    Ok(LoadedWorkbook {
        model,
        graph,
        import_fallbacks,
        data_tables,
        cached_formula_values,
        structure_hash,
        workbook_hash,
        load_ms,
        graph_build_ms,
        ironcalc_ms,
        data_table_summary,
    })
}

fn collect_circular_formula_values(
    model: &Model<'_>,
    graph: &DependencyGraph,
    allow_cached_fallback: bool,
) -> HashMap<CellId, CachedFormulaValue> {
    let circular_cells = graph.circular_formula_cells().collect::<Vec<_>>();
    let iterative_result = evaluate_iterative_formula_cells(model, circular_cells.iter().copied());
    let mut values = iterative_result
        .values
        .into_iter()
        .map(|(cell, value)| {
            (
                cell,
                CachedFormulaValue {
                    value_kind: value.value_kind,
                    value: value.value,
                },
            )
        })
        .collect::<HashMap<_, _>>();

    if allow_cached_fallback {
        for cell in circular_cells {
            if values.contains_key(&cell) {
                continue;
            }
            let Ok((value_kind, value)) = writeback_value_from_cell_value(
                model.get_cell_value_by_index(cell.sheet, cell.row, cell.column),
            ) else {
                continue;
            };
            values.insert(cell, CachedFormulaValue { value_kind, value });
        }
    }

    values
}

fn apply_changed_cells(
    model: &mut Model<'static>,
    changed_cells: &[ChangedCell],
) -> Result<AppliedChanges, EngineError> {
    let mut applied = AppliedChanges {
        changed_cells: Vec::with_capacity(changed_cells.len()),
        formula_changed: false,
    };

    for change in changed_cells {
        let sheet = sheet_index_by_name(model, &change.sheet_name).ok_or_else(|| {
            EngineError::WorkbookLoad(format!(
                "sheet not found for changed cell: {}",
                change.sheet_name
            ))
        })?;
        let before_formula = cell_has_formula(model, sheet, change.row, change.column);
        let after_formula = change.is_formula || change.input.trim_start().starts_with('=');
        if change.input.is_empty() {
            model
                .cell_clear_contents(sheet, change.row, change.column)
                .map_err(EngineError::Evaluation)?;
        } else {
            model
                .set_user_input(sheet, change.row, change.column, change.input.clone())
                .map_err(EngineError::Evaluation)?;
        }

        if before_formula != after_formula || after_formula {
            applied.formula_changed = true;
        }
        applied
            .changed_cells
            .push(CellId::new(sheet, change.row, change.column));
    }

    Ok(applied)
}

fn sheet_index_by_name(model: &Model<'_>, name: &str) -> Option<u32> {
    model
        .workbook
        .worksheets
        .iter()
        .position(|worksheet| worksheet.name.eq_ignore_ascii_case(name))
        .map(|index| index as u32)
}

fn cell_has_formula(model: &Model<'_>, sheet: u32, row: i32, column: i32) -> bool {
    model
        .workbook
        .worksheet(sheet)
        .ok()
        .and_then(|worksheet| worksheet.cell(row, column))
        .map(|cell| cell.has_formula())
        .unwrap_or(false)
}

fn build_result(
    snapshot: &WorkbookSnapshot,
    plan: &mut CalcPlan,
    workbook_hash: &str,
    graph: &DependencyGraph,
    import_fallbacks: &ImportFallbacks,
    data_tables: &[DataTableRegion],
    model: &Model<'_>,
    cached_formula_values: &HashMap<CellId, CachedFormulaValue>,
    started: Instant,
    timing: TimingBreakdown,
) -> CalcResult {
    let data_table_summary = timing.data_tables.clone();
    let fallback_reasons = merged_fallback_reasons(graph, import_fallbacks, &data_table_summary);
    plan.workbook_hash = workbook_hash.to_string();
    plan.mode = snapshot.mode;
    plan.fallback_reasons = fallback_reasons.clone();

    CalcResult {
        plan: plan.clone(),
        analysis: AnalysisSummary {
            workbook_name: snapshot.workbook_name.clone(),
            ironcalc_can_evaluate: fallback_reasons.is_empty(),
            coverage: coverage_with_import_fallbacks(graph, import_fallbacks, &data_table_summary),
            fallback_reasons,
            fallback_details: merged_fallback_details(graph, import_fallbacks),
        },
        benchmark: benchmark_summary(snapshot, started, timing),
        writeback: build_writeback_plan(
            snapshot,
            model,
            graph,
            import_fallbacks,
            data_tables,
            &data_table_summary,
            cached_formula_values,
        ),
    }
}

fn build_writeback_plan(
    snapshot: &WorkbookSnapshot,
    model: &Model<'_>,
    graph: &DependencyGraph,
    import_fallbacks: &ImportFallbacks,
    data_tables: &[DataTableRegion],
    data_table_summary: &DataTableBenchmarkSummary,
    cached_formula_values: &HashMap<CellId, CachedFormulaValue>,
) -> ExcelWritebackPlan {
    let mut notes = vec![
        "Formulas are preserved; the Excel host must pass a live formula-cache probe before applying returned values.".to_string(),
        "Excel remains the correctness authority for unsupported formulas and restore/rebuild behavior.".to_string(),
    ];

    let fallback_reasons = merged_fallback_reasons(graph, import_fallbacks, data_table_summary);
    if !fallback_reasons.is_empty() {
        let skipped = coverage_with_import_fallbacks(graph, import_fallbacks, data_table_summary)
            .fallback_formula_cells;
        let skipped_reasons = vec![WritebackIssueSummary {
            code: "fallback_regions_present".to_string(),
            count: skipped,
            message:
                "Workbook has fallback regions, so Rust values are not written back for this MVP."
                    .to_string(),
        }];
        notes.push(
            "Live writeback is blocked because the workbook contains fallback regions.".to_string(),
        );
        return writeback_plan(
            WritebackMode::None,
            Vec::new(),
            skipped,
            skipped_reasons,
            notes,
        );
    }

    if snapshot.mode != CalcMode::Recalculate {
        notes.push("Live writeback candidates are returned only for Recalculate runs.".to_string());
        return writeback_plan(WritebackMode::None, Vec::new(), 0, Vec::new(), notes);
    }

    let mut skipped_reasons = Vec::new();
    let cells = collect_formula_writeback_cells(
        model,
        graph,
        data_tables,
        cached_formula_values,
        &mut skipped_reasons,
    );
    let skipped = skipped_reasons.iter().map(|reason| reason.count).sum();
    let mode = if cells.is_empty() {
        notes.push(
            "No supported scalar formula values were available for live writeback.".to_string(),
        );
        WritebackMode::None
    } else {
        notes.push(format!(
            "Rust produced {} formula-cache candidate cells for the Excel host.",
            cells.len()
        ));
        WritebackMode::LiveFormulaCache
    };

    writeback_plan(mode, cells, skipped, skipped_reasons, notes)
}

fn writeback_plan(
    mode: WritebackMode,
    cells: Vec<FormulaWritebackCell>,
    skipped: usize,
    skipped_reasons: Vec<WritebackIssueSummary>,
    notes: Vec<String>,
) -> ExcelWritebackPlan {
    ExcelWritebackPlan {
        preserve_formulas: true,
        value_cells_to_update: cells.len(),
        mode,
        cells,
        attempted: 0,
        written: 0,
        skipped,
        failed: 0,
        skipped_reasons,
        failed_samples: Vec::new(),
        notes,
    }
}

fn collect_formula_writeback_cells(
    model: &Model<'_>,
    graph: &DependencyGraph,
    data_tables: &[DataTableRegion],
    cached_formula_values: &HashMap<CellId, CachedFormulaValue>,
    skipped_reasons: &mut Vec<WritebackIssueSummary>,
) -> Vec<FormulaWritebackCell> {
    let mut cells = Vec::new();

    for cell in graph.supported_formula_cells() {
        if is_data_table_output_cell(model, data_tables, cell) {
            add_writeback_issue(
                skipped_reasons,
                "data_table_output",
                "Data table output cells are skipped by the scalar formula writeback MVP.",
                1,
            );
            continue;
        }

        let Some(sheet_name) = sheet_name_for_cell(model, cell) else {
            add_writeback_issue(
                skipped_reasons,
                "missing_sheet",
                "Formula cell belongs to a sheet the host cannot address.",
                1,
            );
            continue;
        };

        let formula = match model.get_cell_formula(cell.sheet, cell.row, cell.column) {
            Ok(Some(formula)) => formula,
            _ => {
                add_writeback_issue(
                    skipped_reasons,
                    "missing_formula",
                    "A candidate cell no longer had formula text available.",
                    1,
                );
                continue;
            }
        };

        if looks_like_multi_cell_formula(&formula) {
            add_writeback_issue(
                skipped_reasons,
                "multi_cell_formula",
                "Array, spill, and known dynamic-array formulas are skipped by the scalar writeback MVP.",
                1,
            );
            continue;
        }

        let (value_kind, value) = if let Some(cached_value) = cached_formula_values.get(&cell) {
            (cached_value.value_kind, cached_value.value.clone())
        } else {
            match writeback_value_from_cell_value(model.get_cell_value_by_index(
                cell.sheet,
                cell.row,
                cell.column,
            )) {
                Ok(value) => value,
                Err((code, message)) => {
                    add_writeback_issue(skipped_reasons, code, message, 1);
                    continue;
                }
            }
        };

        let column = number_to_column(cell.column).unwrap_or_else(|| format!("C{}", cell.column));
        let address = format!("{column}{}", cell.row);
        cells.push(FormulaWritebackCell {
            sheet_name,
            row: cell.row,
            column: cell.column,
            address,
            formula_hash: formula_hash(&formula),
            value_kind,
            value,
        });
    }

    cells
}

fn is_data_table_output_cell(
    model: &Model<'_>,
    data_tables: &[DataTableRegion],
    cell: CellId,
) -> bool {
    let Some(sheet_name) = sheet_name_for_cell(model, cell) else {
        return false;
    };

    data_tables.iter().any(|table| {
        table.sheet_name.eq_ignore_ascii_case(&sheet_name)
            && table.range.contains(cell.row, cell.column)
    })
}

fn sheet_name_for_cell(model: &Model<'_>, cell: CellId) -> Option<String> {
    model
        .workbook
        .worksheets
        .get(cell.sheet as usize)
        .map(|worksheet| worksheet.name.clone())
}

fn writeback_value_from_cell_value(
    value: Result<CellValue, String>,
) -> Result<(FormulaValueKind, serde_json::Value), (&'static str, &'static str)> {
    match value.map_err(|_| {
        (
            "value_read_error",
            "IronCalc could not read the evaluated formula value.",
        )
    })? {
        CellValue::None => Ok((FormulaValueKind::Blank, json!(null))),
        CellValue::Number(value) if value.is_finite() => {
            Ok((FormulaValueKind::Number, json!(value)))
        }
        CellValue::Number(_) => Err((
            "non_finite_number",
            "Formula evaluated to a non-finite numeric value.",
        )),
        CellValue::String(value) if is_excel_error_text(&value) => Err((
            "formula_error",
            "Formula evaluated to an Excel error value.",
        )),
        CellValue::String(value) => Ok((FormulaValueKind::String, json!(value))),
        CellValue::Boolean(value) => Ok((FormulaValueKind::Boolean, json!(value))),
    }
}

fn is_excel_error_text(value: &str) -> bool {
    matches!(
        value.trim().to_ascii_uppercase().as_str(),
        "#NULL!"
            | "#DIV/0!"
            | "#VALUE!"
            | "#REF!"
            | "#NAME?"
            | "#NUM!"
            | "#N/A"
            | "#GETTING_DATA"
            | "#ERROR!"
    )
}

fn looks_like_multi_cell_formula(formula: &str) -> bool {
    let normalized = normalize_formula(formula).to_ascii_uppercase();
    normalized.starts_with('{')
        || normalized.contains('#')
        || [
            "FILTER(",
            "RANDARRAY(",
            "SEQUENCE(",
            "SORT(",
            "SORTBY(",
            "UNIQUE(",
            "TRANSPOSE(",
        ]
        .iter()
        .any(|marker| normalized.contains(marker))
}

fn formula_hash(formula: &str) -> String {
    format!(
        "{:x}",
        Sha256::digest(normalize_formula(formula).as_bytes())
    )
}

fn normalize_formula(formula: &str) -> String {
    formula
        .trim()
        .trim_start_matches('=')
        .trim()
        .replace("\r\n", "\n")
}

fn add_writeback_issue(
    skipped_reasons: &mut Vec<WritebackIssueSummary>,
    code: &str,
    message: &str,
    count: usize,
) {
    if let Some(reason) = skipped_reasons
        .iter_mut()
        .find(|reason| reason.code == code)
    {
        reason.count += count;
        return;
    }

    skipped_reasons.push(WritebackIssueSummary {
        code: code.to_string(),
        count,
        message: message.to_string(),
    });
}

fn benchmark_summary(
    snapshot: &WorkbookSnapshot,
    started: Instant,
    mut timing: TimingBreakdown,
) -> BenchmarkSummary {
    let total_warpspeed_ms = started.elapsed().as_millis();
    let speedup_vs_excel = snapshot.excel_baseline_ms.and_then(|excel_ms| {
        if total_warpspeed_ms == 0 {
            None
        } else {
            Some(excel_ms as f64 / total_warpspeed_ms as f64)
        }
    });

    if timing.result_cache_hit {
        timing.cache_hit_rate = 1.0;
    }

    BenchmarkSummary {
        excel_baseline_ms: snapshot.excel_baseline_ms,
        ironcalc_ms: timing.ironcalc_ms,
        total_warpspeed_ms,
        speedup_vs_excel,
        cache_hit_rate: timing.cache_hit_rate,
        load_ms: timing.load_ms,
        graph_build_ms: timing.graph_build_ms,
        cache_lookup_ms: timing.cache_lookup_ms,
        model_cache_hit: timing.model_cache_hit,
        graph_cache_hit: timing.graph_cache_hit,
        result_cache_hit: timing.result_cache_hit,
        dirty_formula_cells: timing.dirty.dirty_formula_cells,
        planned_reusable_formula_cells: timing.dirty.planned_reusable_formula_cells,
        strategy: timing.strategy,
        data_tables: timing.data_tables,
    }
}

fn load_model_with_import_fallbacks(
    workbook_path: &str,
    workbook_hash: &str,
    locale: &str,
    timezone: &str,
) -> Result<(Model<'static>, ImportFallbacks, Vec<DataTableRegion>), EngineError> {
    match load_from_xlsx(workbook_path, locale, timezone, CACHE_LANGUAGE) {
        Ok(model) => Ok((model, ImportFallbacks::default(), Vec::new())),
        Err(err) if err.to_string().contains("data table formulas") => {
            let Some(sanitized) = sanitize_data_table_formulas(workbook_path, workbook_hash)?
            else {
                return Err(EngineError::WorkbookLoad(err.to_string()));
            };
            let sanitized_path = sanitized.path.to_string_lossy().to_string();
            let loaded = load_from_xlsx(&sanitized_path, locale, timezone, CACHE_LANGUAGE)
                .map_err(|load_err| EngineError::WorkbookLoad(load_err.to_string()));
            remove_sanitized_workbook(&sanitized.path);
            loaded.map(|model| (model, sanitized.fallbacks, sanitized.data_tables))
        }
        Err(err) => Err(EngineError::WorkbookLoad(err.to_string())),
    }
}

fn merged_fallback_reasons(
    graph: &DependencyGraph,
    import_fallbacks: &ImportFallbacks,
    data_table_summary: &DataTableBenchmarkSummary,
) -> Vec<FallbackReason> {
    let mut fallback_reasons = if data_table_fallbacks_are_validated(data_table_summary) {
        Vec::new()
    } else {
        import_fallbacks.fallback_reasons.clone()
    };
    fallback_reasons.extend(graph.fallback_reasons());
    fallback_reasons
}

fn merged_fallback_details(
    graph: &DependencyGraph,
    import_fallbacks: &ImportFallbacks,
) -> Vec<FallbackDetail> {
    let mut fallback_details = import_fallbacks.fallback_details.clone();
    fallback_details.extend(graph.fallback_details());
    fallback_details
}

fn coverage_with_import_fallbacks(
    graph: &DependencyGraph,
    import_fallbacks: &ImportFallbacks,
    data_table_summary: &DataTableBenchmarkSummary,
) -> FormulaCoverage {
    let mut coverage = graph.coverage();
    coverage.formula_cells += import_fallbacks.fallback_formula_cells;
    if !data_table_fallbacks_are_validated(data_table_summary) {
        coverage.fallback_formula_cells += import_fallbacks.fallback_formula_cells;
    }
    coverage
}

fn data_table_fallbacks_are_validated(data_table_summary: &DataTableBenchmarkSummary) -> bool {
    data_table_summary.data_table_count > 0
        && data_table_summary.status == DataTableEvaluationStatus::Validated
        && data_table_summary.data_table_cells == data_table_summary.validated_data_table_cells
        && data_table_summary.mismatched_data_table_cells == 0
        && data_table_summary.unsupported_data_table_cells == 0
}

fn cache_hit_rate(graph: &DependencyGraph, dirty: DirtySummary) -> f64 {
    let formula_count = graph.formula_count();
    if formula_count == 0 {
        1.0
    } else {
        dirty.planned_reusable_formula_cells as f64 / formula_count as f64
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{CalcMode, FormulaCoverage};

    #[test]
    fn formula_coverage_ratio_is_complete_when_no_formulas_exist() {
        let coverage = FormulaCoverage::default();
        assert_eq!(coverage.coverage_ratio(), 1.0);
    }

    #[test]
    fn empty_workbook_path_is_rejected() {
        let snapshot = WorkbookSnapshot {
            workbook_path: String::new(),
            workbook_name: None,
            workbook_id: None,
            mode: CalcMode::Analyze,
            excel_baseline_ms: None,
            force_reload: false,
            changed_cells: Vec::new(),
            evaluate_data_tables: false,
            locale: "en".to_string(),
            timezone: "UTC".to_string(),
            language: "en".to_string(),
        };

        assert!(snapshot.validate().is_err());
    }
}
