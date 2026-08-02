use std::collections::{HashMap, VecDeque};
use std::fs;
use std::sync::Mutex;
use std::time::Instant;

use ironcalc::base::Model;
use ironcalc::import::load_from_xlsx;
use sha2::{Digest, Sha256};

use crate::data_tables::{evaluate_data_tables, summarize_data_tables, DataTableRegion};
use crate::engine::CalcEngine;
use crate::graph::{build_dependency_graph, CellId, DependencyGraph, DirtySummary};
use crate::model::{
    AnalysisSummary, BenchmarkSummary, CalcPlan, CalcResult, CalculationStrategy, ChangedCell,
    DataTableBenchmarkSummary, DataTableEvaluationStatus, EngineError, ExcelWritebackPlan,
    FallbackReason, FormulaCoverage, WorkbookSnapshot,
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
                        result.plan.fallback_reasons =
                            merged_fallback_reasons(&entry.graph, &entry.import_fallbacks);
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
                    let graph_cache_hit = !applied.formula_changed;
                    if applied.formula_changed {
                        let graph_started = Instant::now();
                        entry.graph = build_dependency_graph(&entry.model);
                        entry.structure_hash = entry.graph.structure_hash().to_string();
                        graph_build_ms = graph_started.elapsed().as_millis();
                    }

                    let dirty = entry.graph.dirty_summary(&applied.changed_cells);
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
    let graph = build_dependency_graph(&model);
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
        structure_hash,
        workbook_hash,
        load_ms,
        graph_build_ms,
        ironcalc_ms,
        data_table_summary,
    })
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
    started: Instant,
    timing: TimingBreakdown,
) -> CalcResult {
    let fallback_reasons = merged_fallback_reasons(graph, import_fallbacks);
    plan.workbook_hash = workbook_hash.to_string();
    plan.mode = snapshot.mode;
    plan.fallback_reasons = fallback_reasons.clone();

    CalcResult {
        plan: plan.clone(),
        analysis: AnalysisSummary {
            workbook_name: snapshot.workbook_name.clone(),
            ironcalc_can_evaluate: fallback_reasons.is_empty(),
            coverage: coverage_with_import_fallbacks(graph, import_fallbacks),
            fallback_reasons,
        },
        benchmark: benchmark_summary(snapshot, started, timing),
        writeback: ExcelWritebackPlan {
            preserve_formulas: true,
            value_cells_to_update: 0,
            notes: vec![
                "Prototype mode: formulas are preserved and no cached values are overwritten yet.".to_string(),
                "V1 cache mode reuses workbook state for benchmarking but only skips evaluation on safe no-change cache hits.".to_string(),
            ],
        },
    }
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
) -> Vec<FallbackReason> {
    let mut fallback_reasons = import_fallbacks.fallback_reasons.clone();
    fallback_reasons.extend(graph.fallback_reasons());
    fallback_reasons
}

fn coverage_with_import_fallbacks(
    graph: &DependencyGraph,
    import_fallbacks: &ImportFallbacks,
) -> FormulaCoverage {
    let mut coverage = graph.coverage();
    coverage.formula_cells += import_fallbacks.fallback_formula_cells;
    coverage.fallback_formula_cells += import_fallbacks.fallback_formula_cells;
    coverage
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
