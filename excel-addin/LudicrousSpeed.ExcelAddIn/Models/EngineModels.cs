using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace LudicrousSpeed.ExcelAddIn.Models
{
    internal sealed class WorkbookSnapshot
    {
        [JsonProperty("workbook_path")]
        public string WorkbookPath { get; set; } = "";

        [JsonProperty("workbook_name")]
        public string? WorkbookName { get; set; }

        [JsonProperty("workbook_id")]
        public string? WorkbookId { get; set; }

        [JsonProperty("mode")]
        public string Mode { get; set; } = "analyze";

        [JsonProperty("excel_baseline_ms")]
        public long? ExcelBaselineMs { get; set; }

        [JsonProperty("force_reload")]
        public bool ForceReload { get; set; }

        [JsonProperty("changed_cells")]
        public List<ChangedCell> ChangedCells { get; set; } = new List<ChangedCell>();

        [JsonProperty("evaluate_data_tables")]
        public bool EvaluateDataTables { get; set; }

        [JsonProperty("locale")]
        public string Locale { get; set; } = "en";

        [JsonProperty("timezone")]
        public string Timezone { get; set; } = "UTC";

        [JsonProperty("language")]
        public string Language { get; set; } = "en";

        [JsonProperty("inline_workbook")]
        public InlineWorkbook? InlineWorkbook { get; set; }

        [JsonProperty("data_table_overrides")]
        public List<DataTableOverride> DataTableOverrides { get; set; } = new List<DataTableOverride>();

        /// <summary>
        /// Ask the engine for analysis-only output: per-cell formula hashes
        /// and the full fallback detail list. Roughly 20% of the response on a
        /// large model, and thousands of SHA256 computations, none of it read
        /// unless the report is being written.
        /// </summary>
        [JsonProperty("include_analytics")]
        public bool IncludeAnalytics { get; set; }

        [JsonIgnore]
        public long SnapshotSaveMs { get; set; }

        [JsonIgnore]
        public bool SnapshotSkipped { get; set; }

        [JsonIgnore]
        public string? SheetSignature { get; set; }
    }

    internal sealed class ChangedCell
    {
        [JsonProperty("sheet_name")]
        public string SheetName { get; set; } = "";

        [JsonProperty("row")]
        public int Row { get; set; }

        [JsonProperty("column")]
        public int Column { get; set; }

        [JsonProperty("address")]
        public string Address { get; set; } = "";

        [JsonProperty("input")]
        public string Input { get; set; } = "";

        [JsonProperty("is_formula")]
        public bool IsFormula { get; set; }
    }

    /// <summary>
    /// A full workbook snapshot built directly from a bulk COM read of the
    /// live workbook, as an alternative to SaveCopyAs + re-importing the
    /// saved .xlsx from disk. See InMemoryWorkbookReader.
    /// </summary>
    internal sealed class InlineWorkbook
    {
        [JsonProperty("sheets")]
        public List<InlineSheet> Sheets { get; set; } = new List<InlineSheet>();

        [JsonProperty("defined_names")]
        public List<InlineDefinedName> DefinedNames { get; set; } = new List<InlineDefinedName>();
    }

    internal sealed class InlineSheet
    {
        [JsonProperty("name")]
        public string Name { get; set; } = "";

        [JsonProperty("cells")]
        public List<InlineCell> Cells { get; set; } = new List<InlineCell>();
    }

    internal sealed class InlineCell
    {
        [JsonProperty("row")]
        public int Row { get; set; }

        [JsonProperty("column")]
        public int Column { get; set; }

        [JsonProperty("input")]
        public string Input { get; set; } = "";
    }

    internal sealed class InlineDefinedName
    {
        [JsonProperty("name")]
        public string Name { get; set; } = "";

        [JsonProperty("scope_sheet_name")]
        public string? ScopeSheetName { get; set; }

        [JsonProperty("formula")]
        public string Formula { get; set; } = "";
    }

    internal sealed class HostRunMetrics
    {
        public long? ExcelBaselineMs { get; set; }

        public long SnapshotSaveMs { get; set; }

        public long NativeCallMs { get; set; }

        public long LudicrousEndToEndMs { get; set; }

        public bool SnapshotSkipped { get; set; }

        public long WritebackMs { get; set; }

        public string WritebackStatus { get; set; } = "not_attempted";

        public string? CalculationModeBeforeWriteback { get; set; }

        public string? CalculationModeAfterWriteback { get; set; }

        public int LiveValuesPublished { get; set; }

        /// <summary>
        /// How long writing the report itself took. Worth surfacing: on a
        /// large model the detailed report can cost more than the
        /// calculation, which is not obvious unless it is measured.
        /// </summary>
        public long ReportWriteMs { get; set; }

        public double? EndToEndSpeedupVsExcel
        {
            get
            {
                if (!ExcelBaselineMs.HasValue || LudicrousEndToEndMs == 0)
                {
                    return null;
                }

                return ExcelBaselineMs.Value / (double)LudicrousEndToEndMs;
            }
        }
    }

    internal sealed class EngineResponse
    {
        [JsonProperty("ok")]
        public bool Ok { get; set; }

        [JsonProperty("result")]
        public CalcResult? Result { get; set; }

        [JsonProperty("error")]
        public string? Error { get; set; }

        public static EngineResponse Failed(string error)
        {
            return new EngineResponse { Ok = false, Error = error };
        }
    }

    internal sealed class CalcResult
    {
        [JsonProperty("analysis")]
        public AnalysisSummary Analysis { get; set; } = new AnalysisSummary();

        [JsonProperty("benchmark")]
        public BenchmarkSummary Benchmark { get; set; } = new BenchmarkSummary();

        [JsonProperty("writeback")]
        public ExcelWritebackPlan Writeback { get; set; } = new ExcelWritebackPlan();
    }

    internal sealed class AnalysisSummary
    {
        [JsonProperty("workbook_name")]
        public string? WorkbookName { get; set; }

        [JsonProperty("coverage")]
        public FormulaCoverage Coverage { get; set; } = new FormulaCoverage();

        [JsonProperty("fallback_reasons")]
        public List<FallbackReason> FallbackReasons { get; set; } = new List<FallbackReason>();

        [JsonProperty("fallback_details")]
        public List<FallbackDetail> FallbackDetails { get; set; } = new List<FallbackDetail>();

        [JsonProperty("ironcalc_can_evaluate")]
        public bool IronCalcCanEvaluate { get; set; }
    }

    internal sealed class FormulaCoverage
    {
        [JsonProperty("formula_cells")]
        public int FormulaCells { get; set; }

        [JsonProperty("supported_formula_cells")]
        public int SupportedFormulaCells { get; set; }

        [JsonProperty("fallback_formula_cells")]
        public int FallbackFormulaCells { get; set; }
    }

    internal sealed class FallbackReason
    {
        [JsonProperty("code")]
        public string Code { get; set; } = "";

        [JsonProperty("message")]
        public string Message { get; set; } = "";

        [JsonProperty("location")]
        public string? Location { get; set; }
    }

    internal sealed class FallbackDetail
    {
        [JsonProperty("code")]
        public string Code { get; set; } = "";

        [JsonProperty("message")]
        public string Message { get; set; } = "";

        [JsonProperty("location")]
        public string? Location { get; set; }

        [JsonProperty("formula")]
        public string? Formula { get; set; }

        [JsonProperty("circular_component")]
        public int? CircularComponent { get; set; }

        [JsonProperty("circular_component_size")]
        public int? CircularComponentSize { get; set; }
    }

    internal sealed class BenchmarkSummary
    {
        [JsonProperty("excel_baseline_ms")]
        public long? ExcelBaselineMs { get; set; }

        [JsonProperty("ironcalc_ms")]
        public long IronCalcMs { get; set; }

        [JsonProperty("total_ludicrous_ms")]
        public long TotalLudicrousMs { get; set; }

        [JsonProperty("speedup_vs_excel")]
        public double? SpeedupVsExcel { get; set; }

        [JsonProperty("cache_hit_rate")]
        public double CacheHitRate { get; set; }

        [JsonProperty("load_ms")]
        public long LoadMs { get; set; }

        [JsonProperty("graph_build_ms")]
        public long GraphBuildMs { get; set; }

        [JsonProperty("cache_lookup_ms")]
        public long CacheLookupMs { get; set; }

        [JsonProperty("model_cache_hit")]
        public bool ModelCacheHit { get; set; }

        [JsonProperty("graph_cache_hit")]
        public bool GraphCacheHit { get; set; }

        [JsonProperty("result_cache_hit")]
        public bool ResultCacheHit { get; set; }

        [JsonProperty("dirty_formula_cells")]
        public int DirtyFormulaCells { get; set; }

        [JsonProperty("planned_reusable_formula_cells")]
        public int PlannedReusableFormulaCells { get; set; }

        [JsonProperty("strategy")]
        public string Strategy { get; set; } = "";

        [JsonProperty("data_tables")]
        public DataTableBenchmarkSummary DataTables { get; set; } = new DataTableBenchmarkSummary();
    }

    internal sealed class DataTableBenchmarkSummary
    {
        [JsonProperty("data_table_count")]
        public int DataTableCount { get; set; }

        [JsonProperty("data_table_cells")]
        public int DataTableCells { get; set; }

        [JsonProperty("dirty_data_tables")]
        public int DirtyDataTables { get; set; }

        [JsonProperty("reused_data_table_cells")]
        public int ReusedDataTableCells { get; set; }

        [JsonProperty("evaluated_data_table_cells")]
        public int EvaluatedDataTableCells { get; set; }

        [JsonProperty("validated_data_table_cells")]
        public int ValidatedDataTableCells { get; set; }

        [JsonProperty("mismatched_data_table_cells")]
        public int MismatchedDataTableCells { get; set; }

        [JsonProperty("stale_cache_data_table_cells")]
        public int StaleCacheDataTableCells { get; set; }

        [JsonProperty("unsupported_data_table_cells")]
        public int UnsupportedDataTableCells { get; set; }

        [JsonProperty("data_table_eval_ms")]
        public long DataTableEvalMs { get; set; }

        [JsonProperty("data_table_parallelism")]
        public int DataTableParallelism { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; } = "";

        [JsonProperty("diagnostics")]
        public List<DataTableDiagnostic> Diagnostics { get; set; } = new List<DataTableDiagnostic>();

        [JsonProperty("regions")]
        public List<DataTableRegionInfo> Regions { get; set; } = new List<DataTableRegionInfo>();
    }

    /// <summary>
    /// A native Excel data table's definition, enough to replace it with live
    /// cells and put it back. ColumnInputCell/RowInputCell are the two
    /// arguments Excel's Data Table dialog takes.
    /// </summary>
    /// <summary>
    /// A data table the host has replaced with live cells. Converting deletes
    /// the {=TABLE()} marker the engine discovers tables from, so these
    /// definitions must be replayed on every snapshot or the engine will find
    /// no table there and the live cells will stop updating.
    /// </summary>
    internal sealed class DataTableOverride
    {
        [JsonProperty("sheet_name")]
        public string SheetName { get; set; } = "";

        [JsonProperty("range_address")]
        public string RangeAddress { get; set; } = "";

        [JsonProperty("anchor_address")]
        public string AnchorAddress { get; set; } = "";

        [JsonProperty("column_input_cell")]
        public string? ColumnInputCell { get; set; }

        [JsonProperty("row_input_cell")]
        public string? RowInputCell { get; set; }

        [JsonProperty("is_two_dimensional")]
        public bool IsTwoDimensional { get; set; }

        [JsonProperty("dtr")]
        public bool Dtr { get; set; }
    }

    internal sealed class DataTableRegionInfo
    {
        [JsonProperty("table_id")]
        public string TableId { get; set; } = "";

        [JsonProperty("sheet_name")]
        public string SheetName { get; set; } = "";

        [JsonProperty("range_address")]
        public string RangeAddress { get; set; } = "";

        [JsonProperty("formula_cell")]
        public string? FormulaCell { get; set; }

        [JsonProperty("column_input_cell")]
        public string? ColumnInputCell { get; set; }

        [JsonProperty("row_input_cell")]
        public string? RowInputCell { get; set; }

        [JsonProperty("is_two_dimensional")]
        public bool IsTwoDimensional { get; set; }

        [JsonProperty("dtr")]
        public bool Dtr { get; set; }

        [JsonProperty("kernel_eligible")]
        public bool KernelEligible { get; set; }

        [JsonProperty("cell_count")]
        public int CellCount { get; set; }
    }

    internal sealed class DataTableDiagnostic
    {
        [JsonProperty("code")]
        public string Code { get; set; } = "";

        [JsonProperty("message")]
        public string Message { get; set; } = "";

        [JsonProperty("table_id")]
        public string TableId { get; set; } = "";

        [JsonProperty("sheet_name")]
        public string SheetName { get; set; } = "";

        [JsonProperty("range_address")]
        public string RangeAddress { get; set; } = "";

        [JsonProperty("formula_cell")]
        public string? FormulaCell { get; set; }

        [JsonProperty("formula")]
        public string? Formula { get; set; }

        [JsonProperty("affected_cells")]
        public int AffectedCells { get; set; }
    }

    internal sealed class ExcelWritebackPlan
    {
        [JsonProperty("preserve_formulas")]
        public bool PreserveFormulas { get; set; }

        [JsonProperty("value_cells_to_update")]
        public int ValueCellsToUpdate { get; set; }

        [JsonProperty("mode")]
        public string Mode { get; set; } = "none";

        [JsonProperty("cells")]
        public List<FormulaWritebackCell> Cells { get; set; } = new List<FormulaWritebackCell>();

        [JsonProperty("data_table_cells")]
        public List<DataTableCellValue> DataTableCells { get; set; } = new List<DataTableCellValue>();

        [JsonProperty("skipped")]
        public int Skipped { get; set; }

        [JsonProperty("skipped_reasons")]
        public List<WritebackIssueSummary> SkippedReasons { get; set; } = new List<WritebackIssueSummary>();

        [JsonProperty("notes")]
        public List<string> Notes { get; set; } = new List<string>();
    }

    internal sealed class FormulaWritebackCell
    {
        [JsonProperty("sheet_name")]
        public string SheetName { get; set; } = "";

        [JsonProperty("row")]
        public int Row { get; set; }

        [JsonProperty("column")]
        public int Column { get; set; }

        [JsonProperty("address")]
        public string Address { get; set; } = "";

        [JsonProperty("value_kind")]
        public string ValueKind { get; set; } = "";

        [JsonProperty("value")]
        public JToken? Value { get; set; }
    }

    /// <summary>
    /// A value the data-table kernel computed for a native Excel data table
    /// output cell. Separate from FormulaWritebackCell because these cells
    /// carry no formula of their own -- they're members of an Excel-generated
    /// {=TABLE()} array -- and because they are the highest-value cells to
    /// drive from LudicrousSpeed: Excel re-evaluates the table's source formula
    /// cone once per scenario cell.
    /// </summary>
    internal sealed class DataTableCellValue
    {
        [JsonProperty("sheet_name")]
        public string SheetName { get; set; } = "";

        [JsonProperty("row")]
        public int Row { get; set; }

        [JsonProperty("column")]
        public int Column { get; set; }

        [JsonProperty("address")]
        public string Address { get; set; } = "";

        [JsonProperty("table_id")]
        public string TableId { get; set; } = "";

        /// <summary>
        /// Null when there was no Excel value to compare against, because the
        /// table is already driven by LudicrousSpeed and its "cached" body holds
        /// LudicrousSpeed's own previous output.
        /// </summary>
        [JsonProperty("matched_excel_cache")]
        public bool? MatchedExcelCache { get; set; }

        [JsonProperty("value_kind")]
        public string ValueKind { get; set; } = "";

        [JsonProperty("value")]
        public JToken? Value { get; set; }
    }

    internal sealed class WritebackIssueSummary
    {
        [JsonProperty("code")]
        public string Code { get; set; } = "";

        [JsonProperty("count")]
        public int Count { get; set; }

        [JsonProperty("message")]
        public string Message { get; set; } = "";
    }

    internal sealed class WritebackCellFailure
    {
        [JsonProperty("sheet_name")]
        public string SheetName { get; set; } = "";

        [JsonProperty("address")]
        public string Address { get; set; } = "";

        [JsonProperty("message")]
        public string Message { get; set; } = "";
    }
}
