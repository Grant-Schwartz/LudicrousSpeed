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

                cell.Formula = "=1+1";

                var formulaBefore = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";
                const double injectedValue = 12345.6789;

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

                return LiveFormulaCacheProbeResult.Unsupported(
                    "No supported live formula-cache setter is available. COM Value2 does not satisfy the contract because it replaces the formula instead of updating only the cached result.");
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
