using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace WarpSpeed.ExcelAddIn.Services
{
    internal sealed class LiveFormulaCacheProbe
    {
        private LiveFormulaCacheProbeResult? cachedResult;

        public LiveFormulaCacheProbeResult GetResult()
        {
            if (cachedResult != null)
            {
                return cachedResult;
            }

            cachedResult = RunProbe();
            return cachedResult;
        }

        private static LiveFormulaCacheProbeResult RunProbe()
        {
            Excel.Application excel = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook? scratchWorkbook = null;
            var previousDisplayAlerts = excel.DisplayAlerts;

            try
            {
                excel.DisplayAlerts = false;
                scratchWorkbook = excel.Workbooks.Add();
                Excel.Worksheet sheet = (Excel.Worksheet)scratchWorkbook.Worksheets[1];
                Excel.Range cell = (Excel.Range)sheet.Range["A1"];

                // Calculation mode is per-workbook since Excel 2010, so this
                // freshly-added scratch workbook gets its own default
                // Automatic mode regardless of what the caller's real
                // workbook is set to. xlSet's documented contract is that an
                // injected value persists only until the next recalculation
                // -- under Automatic mode that recalculation can happen
                // essentially immediately, silently overwriting the probe's
                // injected value before it's ever read back. Force Manual
                // here so the probe actually tests the same condition the
                // real writeback runs under (WarpSpeedRibbon's
                // ExcelCalculationGuard already puts the live workbook in
                // Manual mode for the whole run). No need to restore this
                // afterward: the scratch workbook is closed unconditionally
                // in the finally block below and its calculation mode is
                // discarded along with it.
                excel.Calculation = Excel.XlCalculation.xlCalculationManual;

                const double injectedValue = 12345.6789;
                var diagnostics = new List<string>();

                // NOT YET VERIFIED AGAINST LIVE EXCEL: xlSet is documented in
                // the XLL C API SDK as, on a worksheet cell (as opposed to a
                // macro sheet), setting only the cell's calculated/cached
                // value -- the formula itself is untouched, and the injected
                // value is expected to be overwritten by the next real
                // recalculation. That is exactly the formula-preserving
                // contract this probe exists to test, unlike COM Value2
                // (tried below), which always replaces the formula outright.
                // Tried first since it's the theoretically correct
                // mechanism; written and reasoned through carefully but
                // could not be built or run in the environment that
                // authored it (no Windows/Excel available there). See
                // docs/windows-testing.md for the acceptance checklist
                // before trusting this.
                var xlSetResult = TryXlSet(scratchWorkbook, sheet.Name, cell, injectedValue, diagnostics);
                if (xlSetResult != null)
                {
                    return xlSetResult;
                }

                var comValue2Result = TryComValue2(cell, injectedValue, diagnostics);
                if (comValue2Result != null)
                {
                    return comValue2Result;
                }

                return LiveFormulaCacheProbeResult.Unsupported(
                    "No supported live formula-cache setter is available. Neither xlSet nor COM "
                        + "Value2 satisfied the contract of updating only the cached result while "
                        + "preserving the formula. Details: "
                        + string.Join(" | ", diagnostics));
            }
            catch (Exception ex)
            {
                return LiveFormulaCacheProbeResult.Unsupported(
                    "Live formula-cache probe failed: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (scratchWorkbook != null)
                    {
                        scratchWorkbook.Close(false);
                    }
                }
                catch
                {
                    // Scratch workbook cleanup should not hide the probe verdict.
                }

                try
                {
                    excel.DisplayAlerts = previousDisplayAlerts;
                }
                catch
                {
                    // Nothing useful to do if Excel is already tearing down.
                }
            }
        }

        /// <summary>
        /// Attempts the XLL C API's xlSet against <paramref name="cell"/>.
        /// Returns null (rather than an Unsupported result) on any failure
        /// so the caller can fall through to the next candidate mechanism.
        /// </summary>
        private static LiveFormulaCacheProbeResult? TryXlSet(
            Excel.Workbook workbook,
            string sheetName,
            Excel.Range cell,
            double injectedValue,
            List<string> diagnostics)
        {
            try
            {
                cell.Formula = "=1+1";
                var formulaBefore = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";

                var reference = BuildReference(workbook, sheetName, cell.Row, cell.Column, diagnostics);
                if (reference == null)
                {
                    // BuildReference already recorded why.
                    return null;
                }

                object setResult;
                try
                {
                    setResult = XlCall.Excel(XlCall.xlSet, reference, injectedValue);
                }
                catch (Exception ex)
                {
                    diagnostics.Add(
                        $"xlSet: XlCall.Excel(xlSet, ...) threw {ex.GetType().Name}: {ex.Message}. "
                        + "This usually means xlSet was called outside a valid XLL macro/function "
                        + "execution context (e.g. from a ribbon callback rather than a UDF or "
                        + "QueueAsMacro-scheduled call).");
                    return null;
                }

                if (setResult is ExcelError excelError)
                {
                    diagnostics.Add($"xlSet: XlCall.Excel(xlSet, ...) returned ExcelError {excelError}.");
                    return null;
                }

                var formulaAfter = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";
                var valueAfter = cell.Value2;
                var formulaPreserved = string.Equals(formulaBefore, formulaAfter, StringComparison.Ordinal);
                var valueInjected = IsSameNumber(valueAfter, injectedValue);

                diagnostics.Add(
                    $"xlSet: call returned normally. formulaBefore=\"{formulaBefore}\" "
                    + $"formulaAfter=\"{formulaAfter}\" valueAfter={valueAfter} "
                    + $"formulaPreserved={formulaPreserved} valueInjected={valueInjected}");

                if (formulaPreserved && valueInjected)
                {
                    return LiveFormulaCacheProbeResult.Supported(
                        "xl_set",
                        "XLL xlSet preserved formula text and updated the displayed value in the probe workbook.");
                }

                return null;
            }
            catch (Exception ex)
            {
                diagnostics.Add($"xlSet: unexpected exception {ex.GetType().Name}: {ex.Message}");
                return null;
            }
        }

        private static LiveFormulaCacheProbeResult? TryComValue2(
            Excel.Range cell,
            double injectedValue,
            List<string> diagnostics)
        {
            cell.Formula = "=1+1";
            var formulaBefore = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";

            cell.Value2 = injectedValue;

            var formulaAfter = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";
            var valueAfter = cell.Value2;
            var formulaPreserved = string.Equals(formulaBefore, formulaAfter, StringComparison.Ordinal);
            var valueInjected = IsSameNumber(valueAfter, injectedValue);

            diagnostics.Add(
                $"com_value2: formulaBefore=\"{formulaBefore}\" formulaAfter=\"{formulaAfter}\" "
                + $"valueAfter={valueAfter} formulaPreserved={formulaPreserved} valueInjected={valueInjected}");

            if (formulaPreserved && valueInjected)
            {
                return LiveFormulaCacheProbeResult.Supported(
                    "com_value2",
                    "COM Value2 preserved formula text and updated the displayed value in the probe workbook.");
            }

            return null;
        }

        /// <summary>
        /// Resolves an Excel-DNA ExcelReference for an arbitrary cell via
        /// the XLL C API, using a workbook-qualified sheet name
        /// ("[Book2]Sheet1") so the lookup can't resolve to a
        /// same-named sheet in a different open workbook.
        /// </summary>
        internal static ExcelReference? BuildReference(
            Excel.Workbook workbook,
            string sheetName,
            int row,
            int column,
            List<string>? diagnostics = null)
        {
            var qualifiedSheetName = $"[{workbook.Name}]{sheetName}";
            object sheetIdResult;
            try
            {
                sheetIdResult = XlCall.Excel(XlCall.xlSheetId, qualifiedSheetName);
            }
            catch (Exception ex)
            {
                diagnostics?.Add(
                    $"xlSet: XlCall.Excel(xlSheetId, \"{qualifiedSheetName}\") threw "
                    + $"{ex.GetType().Name}: {ex.Message}.");
                return null;
            }

            // xlSheetId returns a whole-sheet ExcelReference wrapping the
            // sheet id, not a raw IntPtr -- SheetId is pulled off of it to
            // build a reference to the actual target cell below.
            if (sheetIdResult is not ExcelReference sheetReference)
            {
                diagnostics?.Add(
                    $"xlSet: XlCall.Excel(xlSheetId, \"{qualifiedSheetName}\") returned "
                    + $"{sheetIdResult?.GetType().Name ?? "null"} ({sheetIdResult}) instead of an "
                    + "ExcelReference.");
                return null;
            }

            // The XLL C API is zero-indexed; COM's Row/Column are one-indexed.
            return new ExcelReference(row - 1, row - 1, column - 1, column - 1, sheetReference.SheetId);
        }

        private static bool IsSameNumber(object value, double expected)
        {
            try
            {
                return Math.Abs(Convert.ToDouble(value, CultureInfo.InvariantCulture) - expected) < 0.0000001;
            }
            catch
            {
                return false;
            }
        }
    }

    internal sealed class LiveFormulaCacheProbeResult
    {
        private LiveFormulaCacheProbeResult(bool isSupported, string mechanism, string message)
        {
            IsSupported = isSupported;
            Mechanism = mechanism;
            Message = message;
        }

        public bool IsSupported { get; }

        public string Mechanism { get; }

        public string Message { get; }

        public static LiveFormulaCacheProbeResult Supported(string mechanism, string message)
        {
            return new LiveFormulaCacheProbeResult(true, mechanism, message);
        }

        public static LiveFormulaCacheProbeResult Unsupported(string message)
        {
            return new LiveFormulaCacheProbeResult(false, "none", message);
        }
    }
}
