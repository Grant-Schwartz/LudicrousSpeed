use std::collections::{HashMap, HashSet, VecDeque};
use std::fs;
use std::sync::Mutex;
use std::time::Instant;

use ironcalc::base::{cell::CellValue, expressions::utils::number_to_column, Model};
use ironcalc::import::load_from_xlsx;
use serde_json::json;
use sha2::{Digest, Sha256};

use crate::data_tables::{
    build_data_table_region_from_override, evaluate_data_tables, evaluate_iterative_formula_cells,
    summarize_data_tables, DataTableRegion,
};
use crate::engine::CalcEngine;
use crate::graph::{build_dependency_graph, CellId, DependencyGraph, DirtySummary};
use crate::model::{
    AnalysisSummary, BenchmarkSummary, CalcMode, CalcPlan, CalcResult, CalculationStrategy,
    ChangedCell, DataTableBenchmarkSummary, DataTableCellValue, DataTableEvaluationStatus,
    DataTableOverride, EngineError, ExcelWritebackPlan, FallbackDetail, FallbackReason,
    FormulaCoverage, FormulaValueKind, FormulaWritebackCell, InlineWorkbook, WorkbookSnapshot,
    WritebackIssueSummary, WritebackMode,
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
                        // Nothing was re-evaluated on a no-change run, but the
                        // host still needs current values for every data table
                        // cell it is driving live -- otherwise those cells go
                        // stale the moment a second recalc happens.
                        data_tables.cell_values = entry.data_table_values.clone();
                        result.plan.mode = snapshot.mode;
                        result.plan.workbook_hash = entry.workbook_hash.clone();
                        result.plan.fallback_reasons = merged_fallback_reasons(
                            &entry.graph,
                            &entry.import_fallbacks,
                            &data_tables,
                            snapshot.evaluate_data_tables,
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
                    let mut data_tables = evaluate_data_tables(
                        &entry.model,
                        &entry.data_tables,
                        &applied.changed_cells,
                        snapshot.evaluate_data_tables,
                    );
                    // Only dirty tables were re-evaluated. Fold those results
                    // over the cached set so the host receives current values
                    // for every table, not just the ones that changed.
                    merge_data_table_values(&mut entry.data_table_values, &data_tables.cell_values);
                    data_tables.cell_values = entry.data_table_values.clone();

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
    /// Latest kernel value for every data table output cell, carried across
    /// runs. A warm run only re-evaluates tables whose inputs are dirty, and a
    /// no-change run re-evaluates none at all, so without this the engine
    /// would report values for just those tables (or none) and a host driving
    /// live cells would see the rest silently stop updating.
    data_table_values: Vec<DataTableCellValue>,
    structure_hash: String,
    workbook_hash: String,
    last_result: Option<CalcResult>,
}

/// Replaces cached values for tables that were re-evaluated this run and
/// keeps the rest, so the merged set always covers every table.
fn merge_data_table_values(cached: &mut Vec<DataTableCellValue>, fresh: &[DataTableCellValue]) {
    if fresh.is_empty() {
        return;
    }

    let refreshed = fresh
        .iter()
        .map(|value| value.table_id.as_str())
        .collect::<HashSet<_>>();
    cached.retain(|value| !refreshed.contains(value.table_id.as_str()));
    cached.extend(fresh.iter().cloned());
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
            data_table_values: self.data_table_summary.cell_values.clone(),
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
    let load_started = Instant::now();
    let (mut model, import_fallbacks, data_tables, workbook_hash) =
        if let Some(inline) = snapshot.inline_workbook.as_ref() {
            let workbook_hash = format!(
                "{:x}",
                Sha256::digest(
                    serde_json::to_vec(inline)
                        .map_err(|err| EngineError::WorkbookLoad(err.to_string()))?
                )
            );
            let (model, import_fallbacks) = build_model_from_inline(inline)?;
            // No xlsx to scan for native Excel data-table array formulas in this
            // path yet; any such cells were already sent as plain formula text
            // and will surface as ordinary fallback/parse issues rather than
            // being recognized as data table regions. See InlineWorkbook docs.
            (model, import_fallbacks, Vec::new(), workbook_hash)
        } else {
            let workbook_path = snapshot.workbook_path.trim();
            if workbook_path.is_empty() {
                return Err(EngineError::WorkbookLoad(
                    "cache miss requires workbook_path or inline_workbook".to_string(),
                ));
            }
            let bytes = fs::read(workbook_path)
                .map_err(|err| EngineError::WorkbookLoad(err.to_string()))?;
            let workbook_hash = format!("{:x}", Sha256::digest(&bytes));
            let (model, import_fallbacks, data_tables) = load_model_with_import_fallbacks(
                workbook_path,
                &workbook_hash,
                &snapshot.locale,
                &snapshot.timezone,
            )?;
            (model, import_fallbacks, data_tables, workbook_hash)
        };
    // Tables the host has already converted no longer carry a {=TABLE()}
    // marker to discover, so they arrive as overrides instead. Merge them in,
    // preferring a discovered table if one somehow still exists at the same
    // range (a table that was restored but whose override wasn't cleared).
    let data_tables = merge_data_table_overrides(data_tables, &snapshot.data_table_overrides);
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
        // A changed cell on a sheet the cached model has never seen cannot
        // affect anything the model computes -- nothing in the model can
        // reference it -- so skip it rather than failing the whole run. This
        // happens legitimately whenever a sheet is added after the last
        // snapshot (WarpSpeed's own bookkeeping sheets being the obvious
        // case). When a workbook path is available the caller force-reloads
        // instead, which picks the sheet up properly; when it isn't, a hard
        // error would leave the user with nothing over a cell that provably
        // does not matter.
        let Some(sheet) = sheet_index_by_name(model, &change.sheet_name) else {
            continue;
        };
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
    let fallback_reasons = merged_fallback_reasons(
        graph,
        import_fallbacks,
        &data_table_summary,
        snapshot.evaluate_data_tables,
    );
    plan.workbook_hash = workbook_hash.to_string();
    plan.mode = snapshot.mode;
    plan.fallback_reasons = fallback_reasons.clone();

    CalcResult {
        plan: plan.clone(),
        analysis: AnalysisSummary {
            workbook_name: snapshot.workbook_name.clone(),
            ironcalc_can_evaluate: fallback_reasons.is_empty(),
            coverage: coverage_with_import_fallbacks(
                graph,
                import_fallbacks,
                &data_table_summary,
                snapshot.evaluate_data_tables,
            ),
            fallback_reasons,
            fallback_details: merged_fallback_details(
                graph,
                import_fallbacks,
                &data_table_summary,
                snapshot.evaluate_data_tables,
            ),
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

    if snapshot.mode != CalcMode::Recalculate {
        notes.push("Live writeback candidates are returned only for Recalculate runs.".to_string());
        return writeback_plan(WritebackMode::None, Vec::new(), 0, Vec::new(), notes);
    }

    let fallback_reasons = merged_fallback_reasons(
        graph,
        import_fallbacks,
        data_table_summary,
        snapshot.evaluate_data_tables,
    );
    if !fallback_reasons.is_empty() {
        notes.push(format!(
            "Workbook has {} fallback region(s); formula cells in or downstream of those regions \
             are excluded from writeback below, but cells in other, unaffected regions are still \
             included.",
            fallback_reasons.len()
        ));
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
    let data_table_cells = data_table_summary.cell_values.clone();
    let mode = if cells.is_empty() && data_table_cells.is_empty() {
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

    if !data_table_cells.is_empty() {
        let matched = data_table_cells
            .iter()
            .filter(|cell| cell.matched_excel_cache == Some(true))
            .count();
        let differed = data_table_cells
            .iter()
            .filter(|cell| cell.matched_excel_cache == Some(false))
            .count();
        let uncompared = data_table_cells
            .iter()
            .filter(|cell| cell.matched_excel_cache.is_none())
            .count();
        notes.push(format!(
            "Rust computed {} data table output cells: {} matched Excel's cached values, {} \
             differed (usually a stale Excel cache rather than a kernel error, but verify before \
             driving those cells live), {} had no Excel value to compare against because the table \
             is already driven by WarpSpeed.",
            data_table_cells.len(),
            matched,
            differed,
            uncompared
        ));
    }

    let mut plan = writeback_plan(mode, cells, skipped, skipped_reasons, notes);
    plan.data_table_cells = data_table_cells;
    plan
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
        data_table_cells: Vec::new(),
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

    let writeback_safe_cells = graph.writeback_safe_formula_cells().collect::<Vec<_>>();
    let supported_count = graph.supported_formula_cells().count();
    if supported_count > writeback_safe_cells.len() {
        add_writeback_issue(
            skipped_reasons,
            "downstream_of_fallback",
            "Formula depends, directly or through a range, on a fallback region elsewhere in \
             the workbook, so its value can't be trusted for writeback even though the formula \
             itself uses only supported constructs.",
            supported_count - writeback_safe_cells.len(),
        );
    }

    for cell in writeback_safe_cells {
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

/// Adds host-declared tables to those discovered in the file. A discovered
/// table always wins for the same id: it still has its native array formula,
/// so its body holds real Excel values worth validating against, whereas an
/// override's body holds WarpSpeed's own output.
fn merge_data_table_overrides(
    mut discovered: Vec<DataTableRegion>,
    overrides: &[DataTableOverride],
) -> Vec<DataTableRegion> {
    if overrides.is_empty() {
        return discovered;
    }

    let known = discovered
        .iter()
        .map(|table| table.id.clone())
        .collect::<HashSet<_>>();
    for override_definition in overrides {
        let region = build_data_table_region_from_override(override_definition);
        if !known.contains(&region.id) {
            discovered.push(region);
        }
    }

    discovered
}

fn load_model_with_import_fallbacks(
    workbook_path: &str,
    workbook_hash: &str,
    locale: &str,
    timezone: &str,
) -> Result<(Model<'static>, ImportFallbacks, Vec<DataTableRegion>), EngineError> {
    // Sanitize up front rather than only in response to a load failure. A
    // native data table formula makes IronCalc's import fail outright, so the
    // old error-triggered path caught those -- but a WS.LIVE cell imports
    // perfectly happily as an unknown function, so waiting for an error would
    // never strip it, leaving every converted cell an unsupported_formula
    // fallback that taints everything downstream. sanitize returns None when
    // there is nothing to strip, so an ordinary workbook is unaffected beyond
    // a cheap string scan of each sheet.
    if let Some(sanitized) = sanitize_data_table_formulas(workbook_path, workbook_hash)? {
        let sanitized_path = sanitized.path.to_string_lossy().to_string();
        let loaded = load_from_xlsx(&sanitized_path, locale, timezone, CACHE_LANGUAGE)
            .map_err(|load_err| EngineError::WorkbookLoad(load_err.to_string()));
        remove_sanitized_workbook(&sanitized.path);
        return loaded.map(|model| (model, sanitized.fallbacks, sanitized.data_tables));
    }

    load_from_xlsx(workbook_path, locale, timezone, CACHE_LANGUAGE)
        .map(|model| (model, ImportFallbacks::default(), Vec::new()))
        .map_err(|err| EngineError::WorkbookLoad(err.to_string()))
}

/// Builds an IronCalc model directly from cell data already collected by the
/// host (e.g. read in bulk over COM from the live Excel workbook), instead of
/// going through an xlsx file at all. Individual cells or defined names that
/// IronCalc rejects are recorded as fallbacks and skipped rather than failing
/// the whole build, since a single bad cell (for example a native Excel data
/// table array formula sent as plain text, which this path does not yet
/// recognize) shouldn't take down an otherwise-good workbook.
fn build_model_from_inline(
    inline: &InlineWorkbook,
) -> Result<(Model<'static>, ImportFallbacks), EngineError> {
    let mut model = Model::new_empty("workbook", "en", "UTC", CACHE_LANGUAGE)
        .map_err(EngineError::WorkbookLoad)?;
    let mut fallbacks = ImportFallbacks::default();

    for (sheet_index, sheet) in inline.sheets.iter().enumerate() {
        if sheet_index == 0 {
            model
                .rename_sheet_by_index(0, &sheet.name)
                .map_err(EngineError::WorkbookLoad)?;
        } else {
            model
                .add_sheet(&sheet.name)
                .map_err(EngineError::WorkbookLoad)?;
        }
    }

    for (sheet_index, sheet) in inline.sheets.iter().enumerate() {
        for cell in &sheet.cells {
            if let Err(message) = model.set_user_input(
                sheet_index as u32,
                cell.row,
                cell.column,
                cell.input.clone(),
            ) {
                record_inline_fallback(
                    &mut fallbacks,
                    "unsupported_formula",
                    format!("Cell could not be set on the in-memory model: {message}"),
                    Some(format!("{}!R{}C{}", sheet.name, cell.row, cell.column)),
                    Some(cell.input.clone()),
                );
            }
        }
    }

    for defined_name in &inline.defined_names {
        let scope = defined_name
            .scope_sheet_name
            .as_deref()
            .and_then(|name| sheet_index_by_name(&model, name));
        if let Err(message) =
            model.new_defined_name(&defined_name.name, scope, &defined_name.formula)
        {
            record_inline_fallback(
                &mut fallbacks,
                "unsupported_reference",
                format!(
                    "Defined name {} could not be created on the in-memory model: {message}",
                    defined_name.name
                ),
                None,
                Some(defined_name.formula.clone()),
            );
        }
    }

    Ok((model, fallbacks))
}

fn record_inline_fallback(
    fallbacks: &mut ImportFallbacks,
    code: &str,
    message: String,
    location: Option<String>,
    formula: Option<String>,
) {
    fallbacks.fallback_formula_cells += 1;
    fallbacks.fallback_reasons.push(FallbackReason {
        code: code.to_string(),
        message: message.clone(),
        location: location.clone(),
    });
    fallbacks.fallback_details.push(FallbackDetail {
        code: code.to_string(),
        message,
        location,
        formula,
        circular_component: None,
        circular_component_size: None,
    });
}

fn merged_fallback_reasons(
    graph: &DependencyGraph,
    import_fallbacks: &ImportFallbacks,
    data_table_summary: &DataTableBenchmarkSummary,
    evaluate_data_tables: bool,
) -> Vec<FallbackReason> {
    let mut fallback_reasons = import_fallbacks
        .fallback_reasons
        .iter()
        .filter(|reason| {
            !table_fallback_is_resolved(
                &reason.code,
                reason.location.as_deref(),
                data_table_summary,
                evaluate_data_tables,
            )
        })
        .cloned()
        .collect::<Vec<_>>();
    fallback_reasons.extend(graph.fallback_reasons());
    fallback_reasons
}

fn merged_fallback_details(
    graph: &DependencyGraph,
    import_fallbacks: &ImportFallbacks,
    data_table_summary: &DataTableBenchmarkSummary,
    evaluate_data_tables: bool,
) -> Vec<FallbackDetail> {
    let mut fallback_details = import_fallbacks
        .fallback_details
        .iter()
        .filter(|detail| {
            !table_fallback_is_resolved(
                &detail.code,
                detail.location.as_deref(),
                data_table_summary,
                evaluate_data_tables,
            )
        })
        .cloned()
        .collect::<Vec<_>>();
    fallback_details.extend(graph.fallback_details());
    fallback_details
}

fn coverage_with_import_fallbacks(
    graph: &DependencyGraph,
    import_fallbacks: &ImportFallbacks,
    data_table_summary: &DataTableBenchmarkSummary,
    evaluate_data_tables: bool,
) -> FormulaCoverage {
    let mut coverage = graph.coverage();
    coverage.formula_cells += import_fallbacks.fallback_formula_cells;
    let unresolved = import_fallbacks
        .fallback_reasons
        .iter()
        .filter(|reason| {
            !table_fallback_is_resolved(
                &reason.code,
                reason.location.as_deref(),
                data_table_summary,
                evaluate_data_tables,
            )
        })
        .count();
    coverage.fallback_formula_cells += unresolved;
    coverage
}

/// A `data_table_formula` import fallback (recorded unconditionally at
/// import time for every native Excel data table cell, before we know
/// whether the table will validate) is resolved -- and can stop being
/// reported as a fallback -- once *that specific table* is confirmed to have
/// validated cleanly: no mismatch, no unsupported-shape diagnostic, nothing.
/// This is deliberately per-table rather than a single workbook-wide gate,
/// so one problematic table (a stale Excel cache, an unsupported function)
/// doesn't mask every other table in the same workbook that validated fine.
///
/// Only trusted when every eligible table was actually (re-)evaluated this
/// run (`reused_data_table_cells == 0`, true for any cold load and any warm
/// run with no dirty-table skips): otherwise "no diagnostic this round"
/// could just mean "not dirty, not rechecked" rather than "clean", and a
/// genuinely still-broken table could wrongly look resolved. That case falls
/// back to the previous, safe-but-coarser workbook-wide behavior.
fn table_fallback_is_resolved(
    code: &str,
    location: Option<&str>,
    data_table_summary: &DataTableBenchmarkSummary,
    evaluate_data_tables: bool,
) -> bool {
    if code != "data_table_formula"
        || !evaluate_data_tables
        || data_table_summary.reused_data_table_cells > 0
    {
        return false;
    }
    let Some(location) = location else {
        return false;
    };
    !data_table_summary
        .diagnostics
        .iter()
        .any(|diagnostic| diagnostic.table_id == location)
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
            inline_workbook: None,
            data_table_overrides: Vec::new(),
        };

        assert!(snapshot.validate().is_err());
    }
}
