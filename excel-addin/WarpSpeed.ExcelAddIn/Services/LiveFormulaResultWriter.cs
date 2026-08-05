using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using ExcelDna.Integration;
using Newtonsoft.Json.Linq;
using WarpSpeed.ExcelAddIn.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace WarpSpeed.ExcelAddIn.Services
{
    internal sealed class LiveFormulaResultWriter
    {
        private const int MaxFailedSamples = 20;

        /// <summary>
        /// Off by default: the in-place formula-cache probe is answered (no
        /// mechanism works -- see Apply). Set WARPSPEED_PROBE_INPLACE=1 to
        /// re-run it, e.g. to re-test on a new Excel build.
        /// </summary>
        private static bool InPlaceProbeEnabled =>
            Environment.GetEnvironmentVariable("WARPSPEED_PROBE_INPLACE") == "1";

        private readonly WorkbookChangeTracker changeTracker;
        private readonly LiveFormulaCacheProbe probe = new LiveFormulaCacheProbe();
        private Excel.XlCalculation? previousCalculationMode;

        public LiveFormulaResultWriter(WorkbookChangeTracker changeTracker)
        {
            this.changeTracker = changeTracker;
        }

        public WritebackApplyResult Apply(EngineResponse response, bool allowWriteback)
        {
            if (!response.Ok || response.Result == null)
            {
                return WritebackApplyResult.NotAttempted("engine_error");
            }

            var plan = response.Result.Writeback;
            ResetHostCounts(plan);

            if (!allowWriteback)
            {
                plan.Notes.Add("Host did not attempt live writeback for this command.");
                return WritebackApplyResult.NotAttempted("command_not_writeback");
            }

            if (!string.Equals(plan.Mode, "live_formula_cache", StringComparison.OrdinalIgnoreCase)
                || plan.Cells.Count == 0)
            {
                plan.Notes.Add("Host had no live formula-cache candidate cells to apply.");
                return WritebackApplyResult.NotAttempted("no_candidates");
            }

            if (!InPlaceProbeEnabled)
            {
                // Every candidate mechanism for writing a value into a cell
                // that already holds a formula has been tested against live
                // Excel and failed: Range.Value2 replaces the formula; the
                // XLL C API's xlSet is macro-sheet-only and silently no-ops
                // on a worksheet cell (formula preserved, value unchanged);
                // and Value2-then-restore-formula re-evaluates the cell the
                // instant the formula goes back in. This is an architectural
                // property of Excel -- a formula cell's formula and value are
                // one unit owned by the calc engine -- not a gap in the
                // search, which is why RTD is the mechanism every market-data
                // vendor uses. Engine values now reach the sheet through
                // WS.LIVE cells instead (see LiveValueService), so don't run
                // the probe on every recalc: it spawns a scratch workbook and
                // reports a blocked result we already know the answer to.
                // Set WARPSPEED_PROBE_INPLACE=1 to re-run it deliberately.
                plan.Notes.Add(
                    "In-place writeback into formula cells is not available in Excel (Value2, xlSet, "
                    + "and Value2-then-restore all verified failing). Engine values are delivered to "
                    + "WS.LIVE cells over RTD instead.");
                return WritebackApplyResult.NotAttempted("superseded_by_live_values");
            }

            var probeResult = probe.GetResult();
            plan.Notes.Add("Live formula-cache probe: " + probeResult.Message);
            if (!probeResult.IsSupported)
            {
                AddSkippedReason(
                    plan,
                    "probe_blocked",
                    plan.Cells.Count,
                    "Formula-preserving live cache writes are blocked because no supported host mechanism passed the probe.");
                return WritebackApplyResult.Blocked(probeResult.Message);
            }

            var started = DateTime.UtcNow;
            Excel.Application excel = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook workbook = excel.ActiveWorkbook;
            var calculationBefore = excel.Calculation;
            previousCalculationMode = calculationBefore;

            using (changeTracker.SuspendTracking())
            {
                excel.Calculation = Excel.XlCalculation.xlCalculationManual;

                foreach (var candidate in plan.Cells)
                {
                    ApplyCandidate(workbook, candidate, plan, probeResult.Mechanism);
                }
            }

            var elapsedMs = (long)Math.Max(0, (DateTime.UtcNow - started).TotalMilliseconds);
            var calculationAfter = excel.Calculation;
            if (plan.Failed > 0 || plan.Skipped > 0)
            {
                return WritebackApplyResult.Partial(
                    elapsedMs,
                    calculationBefore,
                    calculationAfter,
                    "Live formula-cache writeback completed with skips or failures.");
            }

            return WritebackApplyResult.Applied(
                elapsedMs,
                calculationBefore,
                calculationAfter,
                "Live formula-cache writeback completed.");
        }

        /// <summary>
        /// Pushes the engine's computed values into <see cref="LiveValueService"/>
        /// so any WS.LIVE cells watching those addresses update in place.
        /// Independent of the Value2/xlSet writeback path above, which cannot
        /// write into cells that already hold formulas -- this is additive and
        /// safe to run even when that path is blocked.
        /// </summary>
        public int PublishLiveValues(EngineResponse response)
        {
            if (!response.Ok || response.Result == null)
            {
                return 0;
            }

            var plan = response.Result.Writeback;
            var published = new List<KeyValuePair<string, object>>(plan.Cells.Count);
            foreach (var candidate in plan.Cells)
            {
                if (!TryConvertValue(candidate, out var value, out _) || value == null)
                {
                    continue;
                }

                published.Add(new KeyValuePair<string, object>(
                    candidate.SheetName + "!" + candidate.Address,
                    value));
            }

            LiveValueService.PublishAll(published);
            return published.Count;
        }

        public string RestoreLastResults()
        {
            Excel.Application excel = (Excel.Application)ExcelDnaUtil.Application;
            excel.CalculateFullRebuild();

            if (previousCalculationMode.HasValue)
            {
                excel.Calculation = previousCalculationMode.Value;
                var restoredMode = previousCalculationMode.Value.ToString();
                previousCalculationMode = null;
                return "Excel full rebuild completed and calculation mode was restored to " + restoredMode + ".";
            }

            return "Excel full rebuild completed. No prior WarpSpeed calculation mode was recorded.";
        }

        private static void ApplyCandidate(
            Excel.Workbook workbook,
            FormulaWritebackCell candidate,
            ExcelWritebackPlan plan,
            string mechanism)
        {
            Excel.Worksheet? worksheet = FindWorksheet(workbook, candidate.SheetName);
            if (worksheet == null)
            {
                AddSkippedReason(
                    plan,
                    "worksheet_missing",
                    1,
                    "A candidate sheet was not found in the active workbook.");
                return;
            }

            Excel.Range cell = (Excel.Range)worksheet.Cells[candidate.Row, candidate.Column];
            var hasFormula = Convert.ToBoolean(cell.HasFormula, CultureInfo.InvariantCulture);
            if (!hasFormula)
            {
                AddSkippedReason(
                    plan,
                    "target_not_formula",
                    1,
                    "A candidate cell no longer contains a formula.");
                return;
            }

            var formula = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";
            if (!string.Equals(FormulaHash(formula), candidate.FormulaHash, StringComparison.OrdinalIgnoreCase))
            {
                AddSkippedReason(
                    plan,
                    "formula_changed",
                    1,
                    "A candidate formula changed after the engine snapshot.");
                return;
            }

            if (!TryConvertValue(candidate, out var comValue, out var conversionError))
            {
                AddSkippedReason(plan, "unsupported_value", 1, conversionError);
                return;
            }

            plan.Attempted++;
            var originalFormula = formula;
            try
            {
                if (string.Equals(mechanism, "xl_set", StringComparison.OrdinalIgnoreCase))
                {
                    var reference = LiveFormulaCacheProbe.BuildReference(
                        workbook, candidate.SheetName, candidate.Row, candidate.Column);
                    if (reference == null)
                    {
                        plan.Failed++;
                        AddFailedSample(plan, candidate, "Could not resolve an XLL reference for this cell.");
                        return;
                    }

                    var setResult = XlCall.Excel(XlCall.xlSet, reference, comValue);
                    if (setResult is ExcelError)
                    {
                        plan.Failed++;
                        AddFailedSample(plan, candidate, "xlSet returned an error for this cell.");
                        return;
                    }
                }
                else if (string.Equals(mechanism, "value2_then_restore_formula", StringComparison.OrdinalIgnoreCase))
                {
                    // Value2 always destroys the formula outright; restoring
                    // it immediately afterward is the expected shape of this
                    // mechanism, not a failure signal the way an unexpected
                    // formula change is for the plain com_value2 path below.
                    cell.Value2 = comValue;
                    cell.Formula = originalFormula;
                }
                else
                {
                    cell.Value2 = comValue;
                }

                var formulaAfter = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";
                if (!string.Equals(originalFormula, formulaAfter, StringComparison.Ordinal))
                {
                    cell.Formula = originalFormula;
                    plan.Failed++;
                    AddFailedSample(
                        plan,
                        candidate,
                        "Host cache write mechanism replaced the formula; formula text was restored.");
                    return;
                }

                if (!CellValueMatches(cell.Value2, comValue, candidate.ValueKind))
                {
                    plan.Failed++;
                    AddFailedSample(plan, candidate, "Displayed value did not match the Rust result after writeback.");
                    return;
                }

                plan.Written++;
            }
            catch (Exception ex)
            {
                plan.Failed++;
                AddFailedSample(plan, candidate, ex.Message);
                try
                {
                    cell.Formula = originalFormula;
                }
                catch
                {
                    // Best effort restore. The failed sample carries the original error.
                }
            }
        }

        private static void ResetHostCounts(ExcelWritebackPlan plan)
        {
            plan.Attempted = 0;
            plan.Written = 0;
            plan.Failed = 0;
            plan.FailedSamples.Clear();
        }

        private static Excel.Worksheet? FindWorksheet(Excel.Workbook workbook, string sheetName)
        {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                if (string.Equals(
                    Convert.ToString(worksheet.Name, CultureInfo.InvariantCulture),
                    sheetName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }

            return null;
        }

        private static bool TryConvertValue(
            FormulaWritebackCell candidate,
            out object? value,
            out string error)
        {
            value = null;
            error = "";
            var kind = candidate.ValueKind.Trim().ToLowerInvariant();
            JToken? token = candidate.Value;

            try
            {
                switch (kind)
                {
                    case "number":
                        value = token?.ToObject<double>() ?? 0.0;
                        return true;
                    case "string":
                        value = token?.ToObject<string>() ?? "";
                        return true;
                    case "boolean":
                        value = token?.ToObject<bool>() ?? false;
                        return true;
                    case "blank":
                        error = "Blank formula-cache values are skipped because clearing a live formula cell is not safe through COM.";
                        return false;
                    default:
                        error = "Unsupported formula value kind: " + candidate.ValueKind;
                        return false;
                }
            }
            catch (Exception ex)
            {
                error = "Could not convert formula result value: " + ex.Message;
                return false;
            }
        }

        private static bool CellValueMatches(object actual, object? expected, string valueKind)
        {
            var kind = valueKind.Trim().ToLowerInvariant();
            try
            {
                if (kind == "number")
                {
                    return Math.Abs(
                        Convert.ToDouble(actual, CultureInfo.InvariantCulture)
                        - Convert.ToDouble(expected, CultureInfo.InvariantCulture)) < 0.0000001;
                }

                if (kind == "boolean")
                {
                    return Convert.ToBoolean(actual, CultureInfo.InvariantCulture)
                        == Convert.ToBoolean(expected, CultureInfo.InvariantCulture);
                }

                return string.Equals(
                    Convert.ToString(actual, CultureInfo.InvariantCulture),
                    Convert.ToString(expected, CultureInfo.InvariantCulture),
                    StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
        }

        private static string FormulaHash(string formula)
        {
            using (var sha = SHA256.Create())
            {
                var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(NormalizeFormula(formula)));
                var builder = new StringBuilder(hash.Length * 2);
                foreach (var value in hash)
                {
                    builder.Append(value.ToString("x2", CultureInfo.InvariantCulture));
                }

                return builder.ToString();
            }
        }

        private static string NormalizeFormula(string formula)
        {
            var normalized = formula.Trim();
            if (normalized.StartsWith("=", StringComparison.Ordinal))
            {
                normalized = normalized.Substring(1).Trim();
            }

            return normalized.Replace("\r\n", "\n");
        }

        private static void AddSkippedReason(
            ExcelWritebackPlan plan,
            string code,
            int count,
            string message)
        {
            plan.Skipped += count;
            var existing = plan.SkippedReasons.Find(reason =>
                string.Equals(reason.Code, code, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.Count += count;
                return;
            }

            plan.SkippedReasons.Add(new WritebackIssueSummary
            {
                Code = code,
                Count = count,
                Message = message,
            });
        }

        private static void AddFailedSample(
            ExcelWritebackPlan plan,
            FormulaWritebackCell candidate,
            string message)
        {
            if (plan.FailedSamples.Count >= MaxFailedSamples)
            {
                return;
            }

            plan.FailedSamples.Add(new WritebackCellFailure
            {
                SheetName = candidate.SheetName,
                Address = candidate.Address,
                Message = message,
            });
        }
    }

    internal sealed class WritebackApplyResult
    {
        private WritebackApplyResult(
            string status,
            long writebackMs,
            Excel.XlCalculation? calculationBefore,
            Excel.XlCalculation? calculationAfter,
            string message)
        {
            Status = status;
            WritebackMs = writebackMs;
            CalculationBefore = calculationBefore;
            CalculationAfter = calculationAfter;
            Message = message;
        }

        public string Status { get; }

        public long WritebackMs { get; }

        public Excel.XlCalculation? CalculationBefore { get; }

        public Excel.XlCalculation? CalculationAfter { get; }

        public string Message { get; }

        public static WritebackApplyResult NotAttempted(string status)
        {
            return new WritebackApplyResult(status, 0, null, null, status);
        }

        public static WritebackApplyResult Blocked(string message)
        {
            return new WritebackApplyResult("blocked", 0, null, null, message);
        }

        public static WritebackApplyResult Partial(
            long writebackMs,
            Excel.XlCalculation calculationBefore,
            Excel.XlCalculation calculationAfter,
            string message)
        {
            return new WritebackApplyResult(
                "partial",
                writebackMs,
                calculationBefore,
                calculationAfter,
                message);
        }

        public static WritebackApplyResult Applied(
            long writebackMs,
            Excel.XlCalculation calculationBefore,
            Excel.XlCalculation calculationAfter,
            string message)
        {
            return new WritebackApplyResult(
                "applied",
                writebackMs,
                calculationBefore,
                calculationAfter,
                message);
        }
    }
}
