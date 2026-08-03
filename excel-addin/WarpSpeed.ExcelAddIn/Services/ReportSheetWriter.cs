using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using WarpSpeed.ExcelAddIn.Models;

namespace WarpSpeed.ExcelAddIn.Services
{
    internal sealed class ReportSheetWriter
    {
        private const string ReportSheetName = "_WarpSpeed_Report";
        private const int MaxDataTableDiagnostics = 20;
        private const int MaxFallbackSamples = 50;
        private const int MaxWritebackFailures = 20;

        public void Write(EngineResponse response, HostRunMetrics hostMetrics)
        {
            dynamic excel = ExcelDnaUtil.Application;
            dynamic workbook = excel.ActiveWorkbook;
            dynamic sheet = EnsureReportSheet(workbook);

            sheet.Cells.Clear();
            sheet.Range["A1"].Value2 = "WarpSpeed Report";
            sheet.Range["A2"].Value2 = DateTime.Now.ToString("u");

            if (!response.Ok || response.Result == null)
            {
                sheet.Range["A4"].Value2 = "Status";
                sheet.Range["B4"].Value2 = "Error";
                sheet.Range["A5"].Value2 = "Message";
                sheet.Range["B5"].Value2 = response.Error ?? "Unknown error";
                sheet.Columns.AutoFit();
                return;
            }

            var result = response.Result;
            sheet.Range["A4"].Value2 = "Status";
            sheet.Range["B4"].Value2 = "Complete";
            sheet.Range["A6"].Value2 = "Formula cells";
            sheet.Range["B6"].Value2 = result.Analysis.Coverage.FormulaCells;
            sheet.Range["A7"].Value2 = "IronCalc-supported cells";
            sheet.Range["B7"].Value2 = result.Analysis.Coverage.SupportedFormulaCells;
            sheet.Range["A8"].Value2 = "Fallback cells";
            sheet.Range["B8"].Value2 = result.Analysis.Coverage.FallbackFormulaCells;
            sheet.Range["A10"].Value2 = "Excel baseline ms";
            WriteValue(sheet, "B10", hostMetrics.ExcelBaselineMs);
            sheet.Range["A11"].Value2 = "WarpSpeed end-to-end ms";
            sheet.Range["B11"].Value2 = hostMetrics.WarpSpeedEndToEndMs;
            sheet.Range["A12"].Value2 = "End-to-end speedup vs Excel";
            WriteValue(sheet, "B12", hostMetrics.EndToEndSpeedupVsExcel);
            sheet.Range["A13"].Value2 = "Snapshot save ms";
            sheet.Range["B13"].Value2 = hostMetrics.SnapshotSaveMs;
            sheet.Range["A14"].Value2 = "Snapshot skipped";
            sheet.Range["B14"].Value2 = hostMetrics.SnapshotSkipped;
            sheet.Range["A15"].Value2 = "Native call ms";
            sheet.Range["B15"].Value2 = hostMetrics.NativeCallMs;
            sheet.Range["A16"].Value2 = "Rust total ms";
            sheet.Range["B16"].Value2 = result.Benchmark.TotalWarpSpeedMs;
            sheet.Range["A17"].Value2 = "Rust speedup vs Excel";
            WriteValue(sheet, "B17", result.Benchmark.SpeedupVsExcel);
            sheet.Range["A18"].Value2 = "IronCalc evaluate ms";
            sheet.Range["B18"].Value2 = result.Benchmark.IronCalcMs;
            sheet.Range["A19"].Value2 = "Rust load ms";
            sheet.Range["B19"].Value2 = result.Benchmark.LoadMs;
            sheet.Range["A20"].Value2 = "Graph build ms";
            sheet.Range["B20"].Value2 = result.Benchmark.GraphBuildMs;
            sheet.Range["A21"].Value2 = "Cache lookup ms";
            sheet.Range["B21"].Value2 = result.Benchmark.CacheLookupMs;
            sheet.Range["A23"].Value2 = "Strategy";
            sheet.Range["B23"].Value2 = result.Benchmark.Strategy;
            sheet.Range["A24"].Value2 = "Model cache hit";
            sheet.Range["B24"].Value2 = result.Benchmark.ModelCacheHit;
            sheet.Range["A25"].Value2 = "Graph cache hit";
            sheet.Range["B25"].Value2 = result.Benchmark.GraphCacheHit;
            sheet.Range["A26"].Value2 = "Result cache hit";
            sheet.Range["B26"].Value2 = result.Benchmark.ResultCacheHit;
            sheet.Range["A27"].Value2 = "Planned formula reuse rate";
            sheet.Range["B27"].Value2 = result.Benchmark.CacheHitRate;
            sheet.Range["A28"].Value2 = "Dirty formula cells";
            sheet.Range["B28"].Value2 = result.Benchmark.DirtyFormulaCells;
            sheet.Range["A29"].Value2 = "Planned reusable formula cells";
            sheet.Range["B29"].Value2 = result.Benchmark.PlannedReusableFormulaCells;
            sheet.Range["A31"].Value2 = "Data table status";
            sheet.Range["B31"].Value2 = result.Benchmark.DataTables.Status;
            sheet.Range["A32"].Value2 = "Data tables";
            sheet.Range["B32"].Value2 = result.Benchmark.DataTables.DataTableCount;
            sheet.Range["A33"].Value2 = "Data table cells";
            sheet.Range["B33"].Value2 = result.Benchmark.DataTables.DataTableCells;
            sheet.Range["A34"].Value2 = "Dirty data tables";
            sheet.Range["B34"].Value2 = result.Benchmark.DataTables.DirtyDataTables;
            sheet.Range["A35"].Value2 = "Reused data table cells";
            sheet.Range["B35"].Value2 = result.Benchmark.DataTables.ReusedDataTableCells;
            sheet.Range["A36"].Value2 = "Evaluated data table cells";
            sheet.Range["B36"].Value2 = result.Benchmark.DataTables.EvaluatedDataTableCells;
            sheet.Range["A37"].Value2 = "Validated data table cells";
            sheet.Range["B37"].Value2 = result.Benchmark.DataTables.ValidatedDataTableCells;
            sheet.Range["A38"].Value2 = "Mismatched data table cells";
            sheet.Range["B38"].Value2 = result.Benchmark.DataTables.MismatchedDataTableCells;
            sheet.Range["A39"].Value2 = "Unsupported data table cells";
            sheet.Range["B39"].Value2 = result.Benchmark.DataTables.UnsupportedDataTableCells;
            sheet.Range["A40"].Value2 = "Data table eval ms";
            sheet.Range["B40"].Value2 = result.Benchmark.DataTables.DataTableEvalMs;
            sheet.Range["A41"].Value2 = "Data table workers";
            sheet.Range["B41"].Value2 = result.Benchmark.DataTables.DataTableParallelism;

            var row = 43;
            sheet.Range[$"A{row}"].Value2 = "Data table diagnostics";
            sheet.Range[$"B{row}"].Value2 = result.Benchmark.DataTables.Diagnostics.Count;
            if (result.Benchmark.DataTables.Diagnostics.Count > 0)
            {
                row += 2;
                sheet.Range[$"A{row}"].Value2 = "Data table diagnostic code";
                sheet.Range[$"B{row}"].Value2 = "Table";
                sheet.Range[$"C{row}"].Value2 = "Formula cell";
                sheet.Range[$"D{row}"].Value2 = "Affected cells";
                sheet.Range[$"E{row}"].Value2 = "Message";
                sheet.Range[$"F{row}"].Value2 = "Formula";

                var dataTableSampleCount = Math.Min(MaxDataTableDiagnostics, result.Benchmark.DataTables.Diagnostics.Count);
                for (var index = 0; index < dataTableSampleCount; index++)
                {
                    var diagnostic = result.Benchmark.DataTables.Diagnostics[index];
                    row++;
                    sheet.Range[$"A{row}"].Value2 = diagnostic.Code;
                    sheet.Range[$"B{row}"].Value2 = diagnostic.TableId;
                    WriteValue(sheet, $"C{row}", diagnostic.FormulaCell);
                    sheet.Range[$"D{row}"].Value2 = diagnostic.AffectedCells;
                    sheet.Range[$"E{row}"].Value2 = diagnostic.Message;
                    WriteValue(sheet, $"F{row}", diagnostic.Formula);
                }

                if (result.Benchmark.DataTables.Diagnostics.Count > dataTableSampleCount)
                {
                    row++;
                    sheet.Range[$"A{row}"].Value2 = "Additional data table diagnostics omitted";
                    sheet.Range[$"B{row}"].Value2 = result.Benchmark.DataTables.Diagnostics.Count - dataTableSampleCount;
                }
            }

            row += 2;
            sheet.Range[$"A{row}"].Value2 = "Preserve formulas";
            sheet.Range[$"B{row}"].Value2 = result.Writeback.PreserveFormulas;
            row++;
            sheet.Range[$"A{row}"].Value2 = "Cells to update";
            sheet.Range[$"B{row}"].Value2 = result.Writeback.ValueCellsToUpdate;
            row++;
            sheet.Range[$"A{row}"].Value2 = "Writeback mode";
            sheet.Range[$"B{row}"].Value2 = result.Writeback.Mode;
            row++;
            sheet.Range[$"A{row}"].Value2 = "Host writeback status";
            sheet.Range[$"B{row}"].Value2 = hostMetrics.WritebackStatus;
            row++;
            sheet.Range[$"A{row}"].Value2 = "Host writeback ms";
            sheet.Range[$"B{row}"].Value2 = hostMetrics.WritebackMs;
            row++;
            sheet.Range[$"A{row}"].Value2 = "Calc mode before writeback";
            WriteValue(sheet, $"B{row}", hostMetrics.CalculationModeBeforeWriteback);
            row++;
            sheet.Range[$"A{row}"].Value2 = "Calc mode after writeback";
            WriteValue(sheet, $"B{row}", hostMetrics.CalculationModeAfterWriteback);
            row++;
            sheet.Range[$"A{row}"].Value2 = "Candidate cells";
            sheet.Range[$"B{row}"].Value2 = result.Writeback.Cells.Count;
            row++;
            sheet.Range[$"A{row}"].Value2 = "Attempted cells";
            sheet.Range[$"B{row}"].Value2 = result.Writeback.Attempted;
            row++;
            sheet.Range[$"A{row}"].Value2 = "Written cells";
            sheet.Range[$"B{row}"].Value2 = result.Writeback.Written;
            row++;
            sheet.Range[$"A{row}"].Value2 = "Skipped cells";
            sheet.Range[$"B{row}"].Value2 = result.Writeback.Skipped;
            row++;
            sheet.Range[$"A{row}"].Value2 = "Failed cells";
            sheet.Range[$"B{row}"].Value2 = result.Writeback.Failed;

            row += 2;
            sheet.Range[$"A{row}"].Value2 = "Writeback skipped reasons";
            sheet.Range[$"B{row}"].Value2 = result.Writeback.SkippedReasons.Count;
            if (result.Writeback.SkippedReasons.Count > 0)
            {
                row += 2;
                sheet.Range[$"A{row}"].Value2 = "Writeback skip code";
                sheet.Range[$"B{row}"].Value2 = "Count";
                sheet.Range[$"C{row}"].Value2 = "Message";
                foreach (var reason in result.Writeback.SkippedReasons)
                {
                    row++;
                    sheet.Range[$"A{row}"].Value2 = reason.Code;
                    sheet.Range[$"B{row}"].Value2 = reason.Count;
                    sheet.Range[$"C{row}"].Value2 = reason.Message;
                }
            }

            row += 2;
            sheet.Range[$"A{row}"].Value2 = "Writeback failed samples";
            sheet.Range[$"B{row}"].Value2 = result.Writeback.FailedSamples.Count;
            if (result.Writeback.FailedSamples.Count > 0)
            {
                row += 2;
                sheet.Range[$"A{row}"].Value2 = "Sheet";
                sheet.Range[$"B{row}"].Value2 = "Address";
                sheet.Range[$"C{row}"].Value2 = "Message";
                var writebackFailureCount = Math.Min(MaxWritebackFailures, result.Writeback.FailedSamples.Count);
                for (var index = 0; index < writebackFailureCount; index++)
                {
                    var failure = result.Writeback.FailedSamples[index];
                    row++;
                    sheet.Range[$"A{row}"].Value2 = failure.SheetName;
                    sheet.Range[$"B{row}"].Value2 = failure.Address;
                    sheet.Range[$"C{row}"].Value2 = failure.Message;
                }

                if (result.Writeback.FailedSamples.Count > writebackFailureCount)
                {
                    row++;
                    sheet.Range[$"A{row}"].Value2 = "Additional writeback failures omitted";
                    sheet.Range[$"B{row}"].Value2 = result.Writeback.FailedSamples.Count - writebackFailureCount;
                }
            }

            row += 2;
            sheet.Range[$"A{row}"].Value2 = "Fallback reasons";
            sheet.Range[$"B{row}"].Value2 = result.Analysis.FallbackReasons.Count;

            row += 2;
            sheet.Range[$"A{row}"].Value2 = "Fallback code";
            sheet.Range[$"B{row}"].Value2 = "Count";
            foreach (var item in CountFallbackReasons(result.Analysis.FallbackReasons))
            {
                row++;
                sheet.Range[$"A{row}"].Value2 = item.Key;
                sheet.Range[$"B{row}"].Value2 = item.Value;
            }

            row += 2;
            sheet.Range[$"A{row}"].Value2 = "Sample fallback code";
            sheet.Range[$"B{row}"].Value2 = "Location";
            sheet.Range[$"C{row}"].Value2 = "Message";
            var sampleCount = Math.Min(MaxFallbackSamples, result.Analysis.FallbackReasons.Count);
            for (var index = 0; index < sampleCount; index++)
            {
                var reason = result.Analysis.FallbackReasons[index];
                row++;
                sheet.Range[$"A{row}"].Value2 = reason.Code;
                WriteValue(sheet, $"B{row}", reason.Location);
                sheet.Range[$"C{row}"].Value2 = reason.Message;
            }

            if (result.Analysis.FallbackReasons.Count > sampleCount)
            {
                row++;
                sheet.Range[$"A{row}"].Value2 = "Additional fallbacks omitted";
                sheet.Range[$"B{row}"].Value2 = result.Analysis.FallbackReasons.Count - sampleCount;
            }

            row += 2;
            sheet.Range[$"A{row}"].Value2 = "Writeback notes";
            foreach (var note in result.Writeback.Notes)
            {
                row++;
                sheet.Range[$"A{row}"].Value2 = note;
            }

            sheet.Columns.AutoFit();
        }

        private static List<KeyValuePair<string, int>> CountFallbackReasons(List<FallbackReason> reasons)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var reason in reasons)
            {
                if (!counts.ContainsKey(reason.Code))
                {
                    counts[reason.Code] = 0;
                }

                counts[reason.Code]++;
            }

            var items = new List<KeyValuePair<string, int>>(counts);
            items.Sort((left, right) =>
            {
                var countComparison = right.Value.CompareTo(left.Value);
                return countComparison != 0
                    ? countComparison
                    : string.Compare(left.Key, right.Key, StringComparison.Ordinal);
            });
            return items;
        }

        private static void WriteValue(dynamic sheet, string address, object? value)
        {
            sheet.Range[address].Value2 = value ?? "";
        }

        private static dynamic EnsureReportSheet(dynamic workbook)
        {
            foreach (dynamic worksheet in workbook.Worksheets)
            {
                if (Convert.ToString(worksheet.Name) == ReportSheetName)
                {
                    return worksheet;
                }
            }

            dynamic sheet = workbook.Worksheets.Add();
            sheet.Name = ReportSheetName;
            return sheet;
        }
    }
}
