use serde::{Deserialize, Serialize};
use thiserror::Error;

#[derive(Debug, Error)]
pub enum EngineError {
    #[error("workbook path is empty")]
    EmptyWorkbookPath,
    #[error("failed to load workbook: {0}")]
    WorkbookLoad(String),
    #[error("failed to evaluate workbook: {0}")]
    Evaluation(String),
    #[error("failed to serialize response: {0}")]
    Serialization(String),
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum CalcMode {
    Analyze,
    Recalculate,
    Benchmark,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookSnapshot {
    pub workbook_path: String,
    pub workbook_name: Option<String>,
    #[serde(default)]
    pub workbook_id: Option<String>,
    pub mode: CalcMode,
    pub excel_baseline_ms: Option<u128>,
    #[serde(default)]
    pub force_reload: bool,
    #[serde(default)]
    pub changed_cells: Vec<ChangedCell>,
    #[serde(default)]
    pub evaluate_data_tables: bool,
    #[serde(default = "default_locale")]
    pub locale: String,
    #[serde(default = "default_timezone")]
    pub timezone: String,
    #[serde(default = "default_language")]
    pub language: String,
    /// When present on a cache miss, the engine builds the workbook directly
    /// from this in-memory snapshot instead of reading `workbook_path` from
    /// disk, skipping the SaveCopyAs + xlsx re-import round trip entirely.
    /// `workbook_id` must still be set so the result can be cached/reused by
    /// warm runs the same way a file-backed snapshot is.
    #[serde(default)]
    pub inline_workbook: Option<InlineWorkbook>,
    /// Data tables the host has replaced with LudicrousSpeed live cells.
    ///
    /// Converting a table deletes its `{=TABLE()}` array marker, which is the
    /// only thing the engine can discover a native table from. Without these
    /// the engine would find no table there, compute nothing, and the live
    /// cells would never update again. The host persists these definitions
    /// when it converts and replays them on every snapshot.
    #[serde(default)]
    pub data_table_overrides: Vec<DataTableOverride>,
    /// Whether to produce the full fallback detail list, which only the
    /// host's report sheet reads. Off by default so a plain run does no work
    /// for output nobody looks at.
    #[serde(default)]
    pub include_analytics: bool,
}

/// A data table declared by the host rather than discovered in the file.
/// Mirrors the OOXML `dataTable` formula attributes so it can build the same
/// region a native table would have produced.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct DataTableOverride {
    pub sheet_name: String,
    /// Body range, e.g. `F289:J293`.
    pub range_address: String,
    /// Top-left cell of the body; the OOXML `dataTable` formula lives here.
    #[serde(default)]
    pub anchor_address: String,
    /// OOXML `r1` -- the column-axis input cell.
    #[serde(default)]
    pub column_input_cell: Option<String>,
    /// OOXML `r2` -- the row-axis input cell, two-variable tables only.
    #[serde(default)]
    pub row_input_cell: Option<String>,
    /// OOXML `dt2D`.
    #[serde(default)]
    pub is_two_dimensional: bool,
    /// OOXML `dtr` -- row-oriented when true.
    #[serde(default)]
    pub dtr: bool,
}

impl WorkbookSnapshot {
    pub fn validate(&self) -> Result<(), EngineError> {
        if self.workbook_path.trim().is_empty()
            && self
                .workbook_id
                .as_deref()
                .map(str::trim)
                .unwrap_or_default()
                .is_empty()
        {
            return Err(EngineError::EmptyWorkbookPath);
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InlineWorkbook {
    pub sheets: Vec<InlineSheet>,
    #[serde(default)]
    pub defined_names: Vec<InlineDefinedName>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InlineSheet {
    pub name: String,
    pub cells: Vec<InlineCell>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InlineCell {
    pub row: i32,
    pub column: i32,
    /// Same convention as `ChangedCell::input`: formula text starting with
    /// `=`, or the literal value text otherwise.
    pub input: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InlineDefinedName {
    pub name: String,
    /// `None` for a workbook-scoped name; `Some(sheet name)` for a
    /// sheet-scoped one. Resolved to an index against `sheets` when building
    /// the model.
    #[serde(default)]
    pub scope_sheet_name: Option<String>,
    pub formula: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct ChangedCell {
    pub sheet_name: String,
    pub row: i32,
    pub column: i32,
    pub address: String,
    pub input: String,
    pub is_formula: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct FallbackReason {
    pub code: String,
    pub message: String,
    pub location: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct FallbackDetail {
    pub code: String,
    pub message: String,
    pub location: Option<String>,
    pub formula: Option<String>,
    pub circular_component: Option<usize>,
    pub circular_component_size: Option<usize>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default, PartialEq, Eq)]
pub struct FormulaCoverage {
    pub formula_cells: usize,
    pub supported_formula_cells: usize,
    pub fallback_formula_cells: usize,
}

impl FormulaCoverage {
    pub fn coverage_ratio(&self) -> f64 {
        if self.formula_cells == 0 {
            1.0
        } else {
            self.supported_formula_cells as f64 / self.formula_cells as f64
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct CalcPlan {
    pub workbook_hash: String,
    pub mode: CalcMode,
    pub fallback_reasons: Vec<FallbackReason>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AnalysisSummary {
    pub workbook_name: Option<String>,
    pub coverage: FormulaCoverage,
    pub fallback_reasons: Vec<FallbackReason>,
    pub fallback_details: Vec<FallbackDetail>,
    pub ironcalc_can_evaluate: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BenchmarkSummary {
    pub excel_baseline_ms: Option<u128>,
    pub ironcalc_ms: u128,
    pub total_ludicrous_ms: u128,
    pub speedup_vs_excel: Option<f64>,
    pub cache_hit_rate: f64,
    pub load_ms: u128,
    pub graph_build_ms: u128,
    pub cache_lookup_ms: u128,
    pub model_cache_hit: bool,
    pub graph_cache_hit: bool,
    pub result_cache_hit: bool,
    pub dirty_formula_cells: usize,
    pub planned_reusable_formula_cells: usize,
    pub strategy: CalculationStrategy,
    pub data_tables: DataTableBenchmarkSummary,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum CalculationStrategy {
    ColdFull,
    WarmFullWithDirtyPlan,
    WarmNoopCacheHit,
    ForcedReload,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct DataTableBenchmarkSummary {
    pub data_table_count: usize,
    pub data_table_cells: usize,
    pub dirty_data_tables: usize,
    pub reused_data_table_cells: usize,
    pub evaluated_data_table_cells: usize,
    pub validated_data_table_cells: usize,
    pub mismatched_data_table_cells: usize,
    /// Subset of `mismatched_data_table_cells` where Excel's own cached body
    /// was constant across every scenario in the table, a strong signal that
    /// the cache is stale (never recalculated for these inputs) rather than
    /// that the kernel result is wrong.
    pub stale_cache_data_table_cells: usize,
    pub unsupported_data_table_cells: usize,
    pub data_table_eval_ms: u128,
    pub data_table_parallelism: usize,
    pub status: DataTableEvaluationStatus,
    pub diagnostics: Vec<DataTableDiagnostic>,
    /// Kernel-computed values for each data table output cell, carried here
    /// so the host can drive those cells live instead of Excel re-running the
    /// native table. Empty unless data table evaluation was requested.
    #[serde(default)]
    pub cell_values: Vec<DataTableCellValue>,
    /// Definitions of every native data table found, so a host can convert
    /// them to live cells and restore them later.
    #[serde(default)]
    pub regions: Vec<DataTableRegionInfo>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct DataTableDiagnostic {
    pub code: String,
    pub message: String,
    pub table_id: String,
    pub sheet_name: String,
    pub range_address: String,
    pub formula_cell: Option<String>,
    pub formula: Option<String>,
    pub affected_cells: usize,
}

/// Enough of a native Excel data table's definition for a host to replace it
/// with live cells and put it back afterwards. `row_input_cell` and
/// `column_input_cell` are the two arguments Excel's own Data Table dialog
/// takes, so a restore is a single `Range.Table(rowInput, columnInput)` call.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct DataTableRegionInfo {
    pub table_id: String,
    pub sheet_name: String,
    pub range_address: String,
    pub formula_cell: Option<String>,
    pub column_input_cell: Option<String>,
    pub row_input_cell: Option<String>,
    pub is_two_dimensional: bool,
    /// OOXML `dtr`. For a one-variable table this decides which of Excel's
    /// two Data Table inputs the single axis feeds.
    pub dtr: bool,
    /// True when the kernel can actually evaluate this table's shape. Tables
    /// that aren't eligible must keep their native Excel data table, since
    /// LudicrousSpeed has no values to drive them with.
    pub kernel_eligible: bool,
    pub cell_count: usize,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum DataTableEvaluationStatus {
    None,
    MetadataOnly,
    Reused,
    Validated,
    Mismatch,
    Unsupported,
    Partial,
}

impl Default for DataTableEvaluationStatus {
    fn default() -> Self {
        DataTableEvaluationStatus::None
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExcelWritebackPlan {
    pub preserve_formulas: bool,
    pub value_cells_to_update: usize,
    pub mode: WritebackMode,
    pub cells: Vec<FormulaWritebackCell>,
    /// Values the data-table kernel computed for native Excel data table
    /// output cells. These are deliberately separate from `cells`: after
    /// import sanitizing strips the `{=TABLE()}` array marker, the body cells
    /// carry no formula at all, so they never appear in the dependency graph
    /// and can't be produced by the formula-writeback path. They are also the
    /// highest-value cells to drive from LudicrousSpeed, since Excel re-evaluates
    /// the source formula's whole cone once per scenario.
    #[serde(default)]
    pub data_table_cells: Vec<DataTableCellValue>,
    pub skipped: usize,
    pub skipped_reasons: Vec<WritebackIssueSummary>,
    pub notes: Vec<String>,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum WritebackMode {
    None,
    LiveFormulaCache,
}

impl Default for WritebackMode {
    fn default() -> Self {
        WritebackMode::None
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum FormulaValueKind {
    Blank,
    Number,
    String,
    Boolean,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FormulaWritebackCell {
    pub sheet_name: String,
    pub row: i32,
    pub column: i32,
    pub address: String,
    pub value_kind: FormulaValueKind,
    pub value: serde_json::Value,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct DataTableCellValue {
    pub sheet_name: String,
    pub row: i32,
    pub column: i32,
    pub address: String,
    /// Identifies the owning table (`Sheet!Range`), so a host can convert or
    /// restore one table at a time.
    pub table_id: String,
    /// Whether this value matched Excel's cached one.
    ///
    /// `Some(true)` matched; `Some(false)` disagreed -- which on real models
    /// is usually Excel's cache being stale rather than the kernel being
    /// wrong, but the distinction matters before writing values into a live
    /// model, so it travels with the value rather than only in aggregate
    /// counts. `None` means there was nothing to compare against: the table
    /// was declared by the host because it has already been converted, so the
    /// "cached" values in those cells are LudicrousSpeed's own previous output.
    /// Comparing against them would be self-confirming, not validation.
    pub matched_excel_cache: Option<bool>,
    pub value_kind: FormulaValueKind,
    pub value: serde_json::Value,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WritebackIssueSummary {
    pub code: String,
    pub count: usize,
    pub message: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CalcResult {
    pub plan: CalcPlan,
    pub analysis: AnalysisSummary,
    pub benchmark: BenchmarkSummary,
    pub writeback: ExcelWritebackPlan,
}

fn default_locale() -> String {
    "en".to_string()
}

fn default_timezone() -> String {
    "UTC".to_string()
}

fn default_language() -> String {
    "en".to_string()
}
