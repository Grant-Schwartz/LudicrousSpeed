using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using WarpSpeed.ExcelAddIn.Models;

namespace WarpSpeed.ExcelAddIn.Services
{
    internal sealed class ReportSheetWriter
    {
        private const string ReportSheetName = "_WarpSpeed_Report";
        private const string FallbackSheetName = "_WarpSpeed_Fallbacks";
        private const int MaxDataTableDiagnostics = 20;
        private const int MaxFallbackSamples = 50;
        private const int MaxWritebackFailures = 20;

        /// <summary>
        /// Writes the run report.
        ///
        /// Every value goes out in one bulk range assignment rather than a
        /// cell at a time. The per-cell version issued ~150 separate COM
        /// calls, each crossing the process boundary, which on a large model
        /// cost noticeably more than the calculation it was reporting on.
        ///
        /// <paramref name="detailed"/> is the audit/dev view: per-fallback
        /// samples, data table diagnostics, writeback failures and the
        /// separate fallback-detail sheet. Those scale with the number of
        /// problems found and are what actually make a report slow, so the
        /// default run writes only the fixed summary block.
        /// </summary>
        public void Write(EngineResponse response, HostRunMetrics hostMetrics, bool detailed)
        {
            var started = System.Diagnostics.Stopwatch.StartNew();
            dynamic excel = ExcelDnaUtil.Application;
            dynamic workbook = excel.ActiveWorkbook;
            dynamic sheet = EnsureReportSheet(workbook, ReportSheetName);

            sheet.Cells.Clear();

            var rows = new List<KeyValuePair<string, object?>>
            {
                new KeyValuePair<string, object?>("WarpSpeed Report", DateTime.Now.ToString("u")),
            };

            if (!response.Ok || response.Result == null)
            {
                rows.Add(new KeyValuePair<string, object?>("Status", "Error"));
                rows.Add(new KeyValuePair<string, object?>("Message", response.Error ?? "Unknown error"));
                FlushRows(sheet, rows, detailed);
                if (detailed)
                {
                    WriteFallbackDetails(workbook, new List<FallbackDetail>());
                }

                return;
            }

            var result = response.Result;
            var coverage = result.Analysis.Coverage;
            var benchmark = result.Benchmark;
            var dataTables = benchmark.DataTables;

            void Add(string label, object? value) =>
                rows.Add(new KeyValuePair<string, object?>(label, value));

            Add("Status", "Complete");
            Add("", null);
            Add("Formula cells", coverage.FormulaCells);
            Add("IronCalc-supported cells", coverage.SupportedFormulaCells);
            Add("Fallback cells", coverage.FallbackFormulaCells);
            Add("", null);
            Add("Excel baseline ms (full rebuild, data tables forced on)", hostMetrics.ExcelBaselineMs);
            Add("WarpSpeed end-to-end ms", hostMetrics.WarpSpeedEndToEndMs);
            Add("End-to-end speedup vs Excel", hostMetrics.EndToEndSpeedupVsExcel);
            Add("Snapshot save ms", hostMetrics.SnapshotSaveMs);
            Add("Snapshot skipped", hostMetrics.SnapshotSkipped);
            Add("Native call ms", hostMetrics.NativeCallMs);
            Add("Rust total ms", benchmark.TotalWarpSpeedMs);
            Add("Rust speedup vs Excel", benchmark.SpeedupVsExcel);
            Add("IronCalc evaluate ms", benchmark.IronCalcMs);
            Add("Rust load ms", benchmark.LoadMs);
            Add("Graph build ms", benchmark.GraphBuildMs);
            Add("Cache lookup ms", benchmark.CacheLookupMs);
            Add("", null);
            Add("Strategy", benchmark.Strategy);
            Add("Model cache hit", benchmark.ModelCacheHit);
            Add("Graph cache hit", benchmark.GraphCacheHit);
            Add("Result cache hit", benchmark.ResultCacheHit);
            Add("Planned formula reuse rate", benchmark.CacheHitRate);
            Add("Dirty formula cells", benchmark.DirtyFormulaCells);
            Add("Planned reusable formula cells", benchmark.PlannedReusableFormulaCells);
            Add("", null);
            Add("Data table status", dataTables.Status);
            Add("Data tables", dataTables.DataTableCount);
            Add("Data table cells", dataTables.DataTableCells);
            Add("Dirty data tables", dataTables.DirtyDataTables);
            Add("Reused data table cells", dataTables.ReusedDataTableCells);
            Add("Evaluated data table cells", dataTables.EvaluatedDataTableCells);
            Add("Validated data table cells", dataTables.ValidatedDataTableCells);
            Add("Mismatched data table cells", dataTables.MismatchedDataTableCells);
            Add("Stale-cache data table cells", dataTables.StaleCacheDataTableCells);
            Add("Unsupported data table cells", dataTables.UnsupportedDataTableCells);
            Add("Data table eval ms", dataTables.DataTableEvalMs);
            Add("Data table workers", dataTables.DataTableParallelism);
            Add("", null);
            Add("Live values published for WS.LIVE cells", hostMetrics.LiveValuesPublished);
            Add("Writeback mode", result.Writeback.Mode);
            Add("Candidate cells", result.Writeback.ValueCellsToUpdate);
            Add("Host writeback status", hostMetrics.WritebackStatus);
            Add("Report detail", detailed ? "detailed (audit)" : "summary");

            if (detailed)
            {
                AppendDetailSections(rows, result);
            }

            FlushRows(sheet, rows, detailed);

            if (detailed)
            {
                WriteFallbackDetails(workbook, result.Analysis.FallbackDetails);
            }

            started.Stop();
            hostMetrics.ReportWriteMs = started.ElapsedMilliseconds;
            // Written after the bulk flush because the duration isn't known
            // until then. One extra cell write is not worth restructuring for.
            var reportRow = rows.Count + 1;
            sheet.Range[$"A{reportRow}"].Value2 = "Report write ms";
            sheet.Range[$"B{reportRow}"].Value2 = hostMetrics.ReportWriteMs;
        }

        /// <summary>
        /// The parts that grow with how many problems the run found. These
        /// are the audit trail, and the reason a detailed report costs more.
        /// </summary>
        private static void AppendDetailSections(
            List<KeyValuePair<string, object?>> rows,
            CalcResult result)
        {
            void Add(string label, object? value) =>
                rows.Add(new KeyValuePair<string, object?>(label, value));

            Add("", null);
            Add("Data table diagnostics", result.Benchmark.DataTables.Diagnostics.Count);
            var diagnosticCount = Math.Min(
                MaxDataTableDiagnostics, result.Benchmark.DataTables.Diagnostics.Count);
            for (var index = 0; index < diagnosticCount; index++)
            {
                var diagnostic = result.Benchmark.DataTables.Diagnostics[index];
                Add($"  {diagnostic.Code} {diagnostic.TableId}", diagnostic.Message);
            }

            Add("", null);
            Add("Writeback skipped reasons", result.Writeback.SkippedReasons.Count);
            foreach (var reason in result.Writeback.SkippedReasons)
            {
                Add($"  {reason.Code} ({reason.Count})", reason.Message);
            }

            Add("", null);
            Add("Writeback failures", result.Writeback.Failed);
            var failureCount = Math.Min(MaxWritebackFailures, result.Writeback.FailedSamples.Count);
            for (var index = 0; index < failureCount; index++)
            {
                var failure = result.Writeback.FailedSamples[index];
                Add($"  {failure.SheetName}!{failure.Address}", failure.Message);
            }

            Add("", null);
            Add("Fallback reasons", result.Analysis.FallbackReasons.Count);
            var sampleCount = Math.Min(MaxFallbackSamples, result.Analysis.FallbackReasons.Count);
            for (var index = 0; index < sampleCount; index++)
            {
                var reason = result.Analysis.FallbackReasons[index];
                Add($"  {reason.Code} {reason.Location ?? ""}", reason.Message);
            }

            if (result.Analysis.FallbackReasons.Count > sampleCount)
            {
                Add("  Additional fallbacks omitted", result.Analysis.FallbackReasons.Count - sampleCount);
            }

            Add("", null);
            Add("Writeback notes", result.Writeback.Notes.Count);
            foreach (var note in result.Writeback.Notes)
            {
                Add("  note", note);
            }
        }

        /// <summary>
        /// One COM call for the whole block. AutoFit is skipped outside the
        /// detailed view: it is one of the most expensive things you can ask
        /// a sheet to do and it is purely cosmetic.
        /// </summary>
        private static void FlushRows(
            dynamic sheet,
            List<KeyValuePair<string, object?>> rows,
            bool detailed)
        {
            var values = new object[rows.Count, 2];
            for (var index = 0; index < rows.Count; index++)
            {
                values[index, 0] = rows[index].Key;
                values[index, 1] = rows[index].Value ?? "";
            }

            sheet.Range[$"A1:B{rows.Count}"].Value2 = values;
            sheet.Range["A1"].Font.Bold = true;

            if (detailed)
            {
                sheet.Columns.AutoFit();
            }
            else
            {
                sheet.Range["A:A"].ColumnWidth = 52;
                sheet.Range["B:B"].ColumnWidth = 28;
            }
        }

        private static void WriteFallbackDetails(dynamic workbook, List<FallbackDetail> details)
        {
            dynamic sheet = EnsureReportSheet(workbook, FallbackSheetName);
            sheet.Cells.Clear();

            var values = new object[details.Count + 1, 6];
            values[0, 0] = "Fallback code";
            values[0, 1] = "Location";
            values[0, 2] = "Circular component";
            values[0, 3] = "Component size";
            values[0, 4] = "Message";
            values[0, 5] = "Formula";

            for (var index = 0; index < details.Count; index++)
            {
                var detail = details[index];
                var row = index + 1;
                values[row, 0] = detail.Code;
                values[row, 1] = detail.Location ?? "";
                values[row, 2] = detail.CircularComponent.HasValue
                    ? (object)(detail.CircularComponent.Value + 1)
                    : "";
                values[row, 3] = detail.CircularComponentSize.HasValue
                    ? (object)detail.CircularComponentSize.Value
                    : "";
                values[row, 4] = detail.Message;
                values[row, 5] = detail.Formula ?? "";
            }

            sheet.Range[$"A1:F{details.Count + 1}"].Value2 = values;
            sheet.Columns.AutoFit();
        }

        private static dynamic EnsureReportSheet(dynamic workbook, string sheetName)
        {
            foreach (dynamic worksheet in workbook.Worksheets)
            {
                if (Convert.ToString(worksheet.Name) == sheetName)
                {
                    return worksheet;
                }
            }

            dynamic sheet = workbook.Worksheets.Add();
            sheet.Name = sheetName;
            return sheet;
        }
    }
}
