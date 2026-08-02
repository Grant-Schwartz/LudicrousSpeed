using System.Collections.Generic;
using Newtonsoft.Json;

namespace WarpSpeed.ExcelAddIn.Models
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

    internal sealed class HostRunMetrics
    {
        public long? ExcelBaselineMs { get; set; }

        public long SnapshotSaveMs { get; set; }

        public long NativeCallMs { get; set; }

        public long WarpSpeedEndToEndMs { get; set; }

        public bool SnapshotSkipped { get; set; }

        public double? EndToEndSpeedupVsExcel
        {
            get
            {
                if (!ExcelBaselineMs.HasValue || WarpSpeedEndToEndMs == 0)
                {
                    return null;
                }

                return ExcelBaselineMs.Value / (double)WarpSpeedEndToEndMs;
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

    internal sealed class BenchmarkSummary
    {
        [JsonProperty("excel_baseline_ms")]
        public long? ExcelBaselineMs { get; set; }

        [JsonProperty("ironcalc_ms")]
        public long IronCalcMs { get; set; }

        [JsonProperty("total_warpspeed_ms")]
        public long TotalWarpSpeedMs { get; set; }

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

        [JsonProperty("unsupported_data_table_cells")]
        public int UnsupportedDataTableCells { get; set; }

        [JsonProperty("data_table_eval_ms")]
        public long DataTableEvalMs { get; set; }

        [JsonProperty("data_table_parallelism")]
        public int DataTableParallelism { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; } = "";
    }

    internal sealed class ExcelWritebackPlan
    {
        [JsonProperty("preserve_formulas")]
        public bool PreserveFormulas { get; set; }

        [JsonProperty("value_cells_to_update")]
        public int ValueCellsToUpdate { get; set; }

        [JsonProperty("notes")]
        public List<string> Notes { get; set; } = new List<string>();
    }
}
