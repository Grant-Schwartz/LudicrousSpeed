using System;
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

                const double injectedValue = 12345.6789;

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
                var xlSetResult = TryXlSet(scratchWorkbook, sheet.Name, cell, injectedValue);
                if (xlSetResult != null)
                {
                    return xlSetResult;
                }

                var comValue2Result = TryComValue2(cell, injectedValue);
                if (comValue2Result != null)
                {
                    return comValue2Result;
                }

                return LiveFormulaCacheProbeResult.Unsupported(
                    "No supported live formula-cache setter is available. Neither xlSet nor COM "
                        + "Value2 satisfied the contract of updating only the cached result while "
                        + "preserving the formula.");
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
            double injectedValue)
        {
            try
            {
                cell.Formula = "=1+1";
                var formulaBefore = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";

                var reference = BuildReference(workbook, sheetName, cell.Row, cell.Column);
                if (reference == null)
                {
                    return null;
                }

                var setResult = XlCall.Excel(XlCall.xlSet, reference, injectedValue);
                if (setResult is ExcelError)
                {
                    return null;
                }

                var formulaAfter = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";
                var valueAfter = cell.Value2;
                var formulaPreserved = string.Equals(formulaBefore, formulaAfter, StringComparison.Ordinal);
                var valueInjected = IsSameNumber(valueAfter, injectedValue);

                if (formulaPreserved && valueInjected)
                {
                    return LiveFormulaCacheProbeResult.Supported(
                        "xl_set",
                        "XLL xlSet preserved formula text and updated the displayed value in the probe workbook.");
                }

                return null;
            }
            catch
            {
                // xlSet is only valid to call from certain Excel-DNA
                // execution contexts; a failure here just means this
                // mechanism isn't usable right now, not that the probe
                // itself failed.
                return null;
            }
        }

        private static LiveFormulaCacheProbeResult? TryComValue2(Excel.Range cell, double injectedValue)
        {
            cell.Formula = "=1+1";
            var formulaBefore = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";

            cell.Value2 = injectedValue;

            var formulaAfter = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";
            var valueAfter = cell.Value2;
            var formulaPreserved = string.Equals(formulaBefore, formulaAfter, StringComparison.Ordinal);
            var valueInjected = IsSameNumber(valueAfter, injectedValue);

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
        internal static ExcelReference? BuildReference(Excel.Workbook workbook, string sheetName, int row, int column)
        {
            try
            {
                var qualifiedSheetName = $"[{workbook.Name}]{sheetName}";
                var sheetId = XlCall.Excel(XlCall.xlSheetId, qualifiedSheetName);
                if (sheetId is not IntPtr sheetIdValue)
                {
                    return null;
                }

                // The XLL C API is zero-indexed; COM's Row/Column are one-indexed.
                return new ExcelReference(row - 1, row - 1, column - 1, column - 1, sheetIdValue);
            }
            catch
            {
                return null;
            }
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
