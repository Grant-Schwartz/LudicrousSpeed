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
    pub ironcalc_can_evaluate: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BenchmarkSummary {
    pub excel_baseline_ms: Option<u128>,
    pub ironcalc_ms: u128,
    pub total_warpspeed_ms: u128,
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
    pub unsupported_data_table_cells: usize,
    pub data_table_eval_ms: u128,
    pub data_table_parallelism: usize,
    pub status: DataTableEvaluationStatus,
    pub diagnostics: Vec<DataTableDiagnostic>,
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
    pub attempted: usize,
    pub written: usize,
    pub skipped: usize,
    pub failed: usize,
    pub skipped_reasons: Vec<WritebackIssueSummary>,
    pub failed_samples: Vec<WritebackCellFailure>,
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
    pub formula_hash: String,
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
pub struct WritebackCellFailure {
    pub sheet_name: String,
    pub address: String,
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
