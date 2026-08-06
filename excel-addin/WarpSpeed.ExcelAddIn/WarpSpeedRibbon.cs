using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using WarpSpeed.ExcelAddIn.Interop;
using WarpSpeed.ExcelAddIn.Models;
using WarpSpeed.ExcelAddIn.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace WarpSpeed.ExcelAddIn
{
    [ComVisible(true)]
    public sealed class WarpSpeedRibbon : ExcelRibbon, IExcelAddIn
    {
        private readonly NativeEngineClient engineClient = new NativeEngineClient();
        private readonly WorkbookChangeTracker changeTracker = new WorkbookChangeTracker();
        private readonly WorkbookSnapshotService snapshotService;
        private readonly LiveFormulaResultWriter resultWriter;
        private readonly ReportSheetWriter reportWriter = new ReportSheetWriter();
        private readonly DataTableConverter dataTableConverter;

        /// <summary>
        /// Opt-in only: set WARPSPEED_ASYNC_RUN=1 to run the native engine
        /// call off Excel's UI thread instead of blocking it. Off by default
        /// until verified against live Excel -- see the doc comment on
        /// <see cref="RunAsync"/>.
        /// </summary>
        private static bool AsyncRunEnabled =>
            Environment.GetEnvironmentVariable("WARPSPEED_ASYNC_RUN") == "1";

        /// <summary>
        /// Audit/dev view. Off by default because the detailed sections scale
        /// with how many problems a run found -- per-fallback samples, data
        /// table diagnostics, writeback failures and the separate detail
        /// sheet -- and on a large model that costs more than the calculation
        /// being reported on.
        /// </summary>
        private bool detailedReport;

        public void ToggleDetailedReport(IRibbonControl control, bool pressed)
        {
            detailedReport = pressed;
        }

        public bool GetDetailedReportPressed(IRibbonControl control)
        {
            return detailedReport;
        }

        public WarpSpeedRibbon()
        {
            snapshotService = new WorkbookSnapshotService(changeTracker);
            resultWriter = new LiveFormulaResultWriter(changeTracker);
            dataTableConverter = new DataTableConverter(changeTracker);
        }

        public void AutoOpen()
        {
            changeTracker.Start();
        }

        public void AutoClose()
        {
            changeTracker.Stop();
        }

        public override string GetCustomUI(string ribbonId)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='WarpSpeedTab' label='WarpSpeed'>
        <group id='WarpSpeedCalcGroup' label='Calculation'>
          <button id='AnalyzeWorkbookButton'
                  label='Analyze Workbook'
                  size='large'
                  imageMso='ErrorChecking'
                  onAction='AnalyzeWorkbook'
                  screentip='Analyze Workbook'
                  supertip='Scan the active workbook and report IronCalc coverage and fallback regions.' />
          <button id='RecalculateWorkbookButton'
                  label='Recalculate with WarpSpeed'
                  size='large'
                  imageMso='CalculateNow'
                  onAction='RecalculateWithWarpSpeed'
                  screentip='Recalculate with WarpSpeed'
                  supertip='Evaluate the active workbook with IronCalc first and Excel fallback for unsupported regions.' />
          <button id='BenchmarkWorkbookButton'
                  label='Benchmark'
                  size='large'
                  imageMso='CalculateSheet'
                  onAction='BenchmarkWorkbook'
                  screentip='Benchmark Workbook'
                  supertip='Compare Excel full rebuild timing against the WarpSpeed prototype engine.' />
          <toggleButton id='DetailedReportToggle'
                  label='Detailed Report'
                  imageMso='ErrorChecking'
                  onAction='ToggleDetailedReport'
                  getPressed='GetDetailedReportPressed'
                  screentip='Detailed Report (audit mode)'
                  supertip='Include per-fallback samples, data table diagnostics, writeback failures and the fallback detail sheet. Slower on large models, so off by default.' />
          <button id='RestoreButton'
                  label='Restore Last Results'
                  imageMso='Undo'
                  onAction='RestoreLastResults'
                  screentip='Restore Last Results'
                  supertip='Restore value changes made by the last writeback-capable WarpSpeed run.' />
        </group>
        <group id='WarpSpeedDataTableGroup' label='Data Tables'>
          <button id='ConvertDataTablesButton'
                  label='Convert to Live'
                  size='large'
                  imageMso='Refresh'
                  onAction='ConvertDataTablesToLive'
                  screentip='Convert Data Tables to WarpSpeed Live Cells'
                  supertip='Replace native Excel data tables with WS.LIVE cells driven by the WarpSpeed kernel. Excel stops re-running the table once per scenario; the source formula and axis inputs are left untouched.' />
          <button id='RestoreDataTablesButton'
                  label='Restore Native'
                  size='large'
                  imageMso='Undo'
                  onAction='RestoreNativeDataTables'
                  screentip='Restore Native Excel Data Tables'
                  supertip='Put the original Excel data tables back, using the definitions recorded when they were converted.' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void AnalyzeWorkbook(IRibbonControl control)
        {
            Run("analyze", includeExcelBaseline: false);
        }

        public void RecalculateWithWarpSpeed(IRibbonControl control)
        {
            Run("recalculate", includeExcelBaseline: false);
        }

        public void BenchmarkWorkbook(IRibbonControl control)
        {
            Run("benchmark", includeExcelBaseline: true);
        }

        /// <summary>
        /// Runs the engine to discover this workbook's native data tables,
        /// then replaces each eligible one with WS.LIVE cells. The engine run
        /// is required rather than optional: it both locates the tables (they
        /// are only visible in the OOXML, not through the object model in a
        /// form we can enumerate) and computes the values the live cells will
        /// display.
        /// </summary>
        public void ConvertDataTablesToLive(IRibbonControl control)
        {
            WorkbookSnapshot? snapshot = null;
            // Manual mode for the whole operation, established before the
            // snapshot is taken. Publishing marks every dependent of ~11k RTD
            // values dirty, and SaveCopyAs recalculates a dirty workbook
            // before writing -- under Automatic either one triggers exactly
            // the full recalculation this feature exists to avoid.
            using var calculationGuard = ExcelCalculationGuard.Enter(disableNativeDataTables: true);
            try
            {
                snapshot = snapshotService.Create("recalculate", null);
                snapshot.DataTableOverrides = dataTableConverter.ReadOverrides();
                var response = engineClient.Run(snapshot, out _);
                if (!response.Ok || response.Result == null)
                {
                    ShowError(response.Error ?? "The engine could not analyze this workbook.");
                    return;
                }

                // Publish first: the WS.LIVE formulas written below resolve
                // immediately instead of showing #N/A until the next recalc.
                var published = resultWriter.PublishLiveValues(response);

                var regions = response.Result.Benchmark.DataTables.Regions;
                var confirm = MessageBox.Show(
                    $"Convert {regions.Count} native data table(s) to WarpSpeed live cells?"
                        + Environment.NewLine + Environment.NewLine
                        + "The Excel data tables will be replaced with WS.LIVE formulas. Your source "
                        + "formulas and axis inputs are not modified, and 'Restore Native' puts the "
                        + "original tables back."
                        + Environment.NewLine + Environment.NewLine
                        + $"{published:N0} engine values are available to drive them.",
                    "WarpSpeed",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Question);
                if (confirm != DialogResult.OK)
                {
                    return;
                }

                var result = dataTableConverter.ConvertToLive(regions);
                var message = result.Message;
                if (result.SkippedReasons.Count > 0)
                {
                    message += Environment.NewLine + Environment.NewLine
                        + string.Join(Environment.NewLine, result.SkippedReasons.ToArray());
                }

                MessageBox.Show(message, "WarpSpeed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
            finally
            {
                TryDeleteSnapshot(snapshot);
            }
        }

        public void RestoreNativeDataTables(IRibbonControl control)
        {
            try
            {
                var result = dataTableConverter.RestoreNativeTables();
                var message = result.Message;
                if (result.Errors.Count > 0)
                {
                    message += Environment.NewLine + Environment.NewLine
                        + string.Join(Environment.NewLine, result.Errors.ToArray());
                }

                MessageBox.Show(message, "WarpSpeed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
        }

        public void RestoreLastResults(IRibbonControl control)
        {
            MessageBox.Show(
                resultWriter.RestoreLastResults(),
                "WarpSpeed",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void Run(string mode, bool includeExcelBaseline)
        {
            if (AsyncRunEnabled)
            {
                RunAsync(mode, includeExcelBaseline);
            }
            else
            {
                RunSync(mode, includeExcelBaseline);
            }
        }

        /// <summary>
        /// The original, fully-synchronous path: everything runs on Excel's
        /// UI thread (the ribbon callback thread), including the native
        /// engine call, so Excel is unresponsive for the whole operation.
        /// Kept as the default because it is the one this add-in has actually
        /// been run with.
        /// </summary>
        private void RunSync(string mode, bool includeExcelBaseline)
        {
            WorkbookSnapshot? snapshot = null;
            using var calculationGuard = ExcelCalculationGuard.Enter(disableNativeDataTables: !includeExcelBaseline);
            try
            {
                var excelBaselineMs = MeasureExcelBaseline(includeExcelBaseline);
                var warpspeedStopwatch = Stopwatch.StartNew();
                snapshot = snapshotService.Create(mode, excelBaselineMs);
                snapshot.DataTableOverrides = dataTableConverter.ReadOverrides();
                snapshot.IncludeAnalytics = detailedReport;
                var response = engineClient.Run(snapshot, out var nativeCallMs);
                warpspeedStopwatch.Stop();

                FinishRun(mode, snapshot, response, excelBaselineMs, warpspeedStopwatch.ElapsedMilliseconds, nativeCallMs);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
            finally
            {
                TryDeleteSnapshot(snapshot);
            }
        }

        /// <summary>
        /// NOT YET VERIFIED AGAINST LIVE EXCEL. Written and reasoned through
        /// carefully but never built or run against live Excel (no
        /// Windows/Excel available in the environment that authored it).
        /// Runs only the native engine call (engineClient.Run -- pure FFI,
        /// touches no Excel COM objects) on a background thread, so Excel's
        /// UI thread is free while a large workbook is being evaluated.
        /// Everything that touches the Excel object model -- snapshot
        /// creation, the calculation-mode guard, writeback, the report
        /// sheet, and the completion dialog -- stays on Excel's main thread,
        /// either by running there directly (before the background call) or
        /// via <see cref="ExcelAsyncUtil.QueueAsMacro"/> (after it), which is
        /// Excel-DNA's documented mechanism for scheduling a callback back
        /// onto Excel's main thread from a background thread. Before
        /// enabling this by default: build on Windows, confirm it actually
        /// keeps Excel responsive during a long native call on a large
        /// workbook, and confirm the calculation-mode guard is reliably
        /// restored (including if the native call throws) -- see the
        /// acceptance checklist in docs/windows-testing.md.
        /// </summary>
        private void RunAsync(string mode, bool includeExcelBaseline)
        {
            WorkbookSnapshot? snapshot = null;
            var calculationGuard = ExcelCalculationGuard.Enter(disableNativeDataTables: !includeExcelBaseline);
            try
            {
                var excelBaselineMs = MeasureExcelBaseline(includeExcelBaseline);
                var warpspeedStopwatch = Stopwatch.StartNew();
                snapshot = snapshotService.Create(mode, excelBaselineMs);
                snapshot.DataTableOverrides = dataTableConverter.ReadOverrides();
                snapshot.IncludeAnalytics = detailedReport;
                var snapshotForBackgroundCall = snapshot;

                SetStatusBar("WarpSpeed is calculating in the background...");

                Task.Run(() =>
                {
                    EngineResponse? response = null;
                    long nativeCallMs = 0;
                    Exception? backgroundError = null;
                    try
                    {
                        response = engineClient.Run(snapshotForBackgroundCall, out nativeCallMs);
                    }
                    catch (Exception ex)
                    {
                        backgroundError = ex;
                    }

                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            if (backgroundError != null)
                            {
                                ShowError(backgroundError.Message);
                                return;
                            }

                            warpspeedStopwatch.Stop();
                            FinishRun(
                                mode,
                                snapshotForBackgroundCall,
                                response!,
                                excelBaselineMs,
                                warpspeedStopwatch.ElapsedMilliseconds,
                                nativeCallMs);
                        }
                        catch (Exception ex)
                        {
                            ShowError(ex.Message);
                        }
                        finally
                        {
                            ClearStatusBar();
                            calculationGuard.Dispose();
                            TryDeleteSnapshot(snapshotForBackgroundCall);
                        }
                    });
                });
            }
            catch (Exception ex)
            {
                ClearStatusBar();
                calculationGuard.Dispose();
                TryDeleteSnapshot(snapshot);
                ShowError(ex.Message);
            }
        }

        /// <summary>
        /// Everything after the native engine call has returned a response:
        /// applying writeback, writing the report sheet, and showing the
        /// completion message. Shared by both the sync and async paths so
        /// enabling WARPSPEED_ASYNC_RUN changes only how/when this runs, not
        /// what it does.
        /// </summary>
        private void FinishRun(
            string mode,
            WorkbookSnapshot snapshot,
            EngineResponse response,
            long? excelBaselineMs,
            long warpSpeedEndToEndMs,
            long nativeCallMs)
        {
            var writebackResult = resultWriter.Apply(
                response,
                string.Equals(mode, "recalculate", StringComparison.OrdinalIgnoreCase));

            // Push engine results to any WS.LIVE cells, then let Excel
            // propagate from those injection points. RTD values land
            // asynchronously, so the recalculation is queued rather than
            // called inline -- QueueAsMacro runs it once Excel has processed
            // the pending updates, otherwise Calculate could run against
            // values that haven't arrived yet and (under Manual mode) never
            // get recalculated again.
            var livePublished = resultWriter.PublishLiveValues(response);
            if (livePublished > 0)
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        dynamic excelApp = ExcelDnaUtil.Application;
                        excelApp.Calculate();
                    }
                    catch
                    {
                        // A failed propagation pass leaves the WS.LIVE cells
                        // themselves correct; the user can press F9.
                    }
                });
            }

            var hostMetrics = new HostRunMetrics
            {
                ExcelBaselineMs = excelBaselineMs,
                SnapshotSaveMs = snapshot.SnapshotSaveMs,
                SnapshotSkipped = snapshot.SnapshotSkipped,
                NativeCallMs = nativeCallMs,
                WarpSpeedEndToEndMs = warpSpeedEndToEndMs,
                WritebackMs = writebackResult.WritebackMs,
                WritebackStatus = writebackResult.Status,
                CalculationModeBeforeWriteback = writebackResult.CalculationBefore?.ToString(),
                CalculationModeAfterWriteback = writebackResult.CalculationAfter?.ToString(),
                LiveValuesPublished = livePublished,
            };

            // No-op unless the Detailed Report toggle is on.
            reportWriter.Write(response, hostMetrics, detailedReport);

            if (!response.Ok)
            {
                ShowError(response.Error);
                return;
            }

            changeTracker.MarkRunSucceeded(snapshot);

            // Success is reported on the status bar rather than in a dialog.
            // A modal popup after every recalculation is the single most
            // intrusive thing an add-in can do, and it also stops the clock on
            // "how fast did that feel". Failures still raise a dialog -- those
            // are worth interrupting for.
            var elapsed = hostMetrics.WarpSpeedEndToEndMs;
            var summary = $"WarpSpeed: {elapsed:N0} ms";
            if (livePublished > 0)
            {
                summary += $", {livePublished:N0} live values published";
            }

            if (detailedReport)
            {
                summary += " - see _WarpSpeed_Report";
            }

            SetStatusBar(summary);
        }

        private static void ShowError(string? message)
        {
            MessageBox.Show(message, "WarpSpeed", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void SetStatusBar(string message)
        {
            try
            {
                dynamic excel = ExcelDnaUtil.Application;
                excel.StatusBar = message;
            }
            catch
            {
                // Status bar text is a courtesy, not load-bearing.
            }
        }

        private static void ClearStatusBar()
        {
            try
            {
                dynamic excel = ExcelDnaUtil.Application;
                excel.StatusBar = false;
            }
            catch
            {
                // Status bar text is a courtesy, not load-bearing.
            }
        }

        private sealed class ExcelCalculationGuard : IDisposable
        {
            private readonly Excel.Application excel;
            private readonly Excel.XlCalculation previousCalculation;
            private readonly bool disableNativeDataTables;
            private bool disposed;

            private ExcelCalculationGuard(Excel.Application excel, bool disableNativeDataTables)
            {
                this.excel = excel;
                this.disableNativeDataTables = disableNativeDataTables;
                previousCalculation = excel.Calculation;

                if (disableNativeDataTables)
                {
                    excel.Calculation = Excel.XlCalculation.xlCalculationManual;
                }
            }

            public static ExcelCalculationGuard Enter(bool disableNativeDataTables)
            {
                return new ExcelCalculationGuard(
                    (Excel.Application)ExcelDnaUtil.Application,
                    disableNativeDataTables);
            }

            public void Dispose()
            {
                if (disposed)
                {
                    return;
                }

                disposed = true;
                if (!disableNativeDataTables)
                {
                    return;
                }

                excel.Calculation = previousCalculation == Excel.XlCalculation.xlCalculationAutomatic
                    ? Excel.XlCalculation.xlCalculationSemiautomatic
                    : previousCalculation;
            }
        }

        /// <summary>
        /// Times a full Excel rebuild, with data tables forced on.
        ///
        /// Calculation mode has to be set to fully Automatic first. Real
        /// models routinely sit in xlCalculationSemiautomatic ("automatic
        /// except data tables") precisely because their tables are too slow to
        /// tolerate -- ALMS_v11 ships that way. Measuring a baseline in that
        /// mode silently skips the most expensive thing in the workbook and
        /// compares WarpSpeed-with-tables against Excel-without-them, which
        /// flatters the result and measures nothing useful.
        /// </summary>
        private static long? MeasureExcelBaseline(bool includeExcelBaseline)
        {
            if (!includeExcelBaseline)
            {
                return null;
            }

            dynamic excel = ExcelDnaUtil.Application;
            var previousCalculation = excel.Calculation;
            try
            {
                excel.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                var stopwatch = Stopwatch.StartNew();
                excel.CalculateFullRebuild();
                stopwatch.Stop();
                return stopwatch.ElapsedMilliseconds;
            }
            finally
            {
                try
                {
                    excel.Calculation = previousCalculation;
                }
                catch
                {
                    // Leaving Excel in Automatic is recoverable; hiding the
                    // measurement behind a restore failure is not.
                }
            }
        }

        private static void TryDeleteSnapshot(WorkbookSnapshot? snapshot)
        {
            if (snapshot == null || string.IsNullOrWhiteSpace(snapshot.WorkbookPath))
            {
                return;
            }

            try
            {
                if (File.Exists(snapshot.WorkbookPath))
                {
                    File.Delete(snapshot.WorkbookPath);
                }
            }
            catch
            {
                // Temp cleanup should not obscure the calculation result.
            }
        }
    }
}
