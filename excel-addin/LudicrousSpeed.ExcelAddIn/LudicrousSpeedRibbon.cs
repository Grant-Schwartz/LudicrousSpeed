using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using LudicrousSpeed.ExcelAddIn.Interop;
using LudicrousSpeed.ExcelAddIn.Models;
using LudicrousSpeed.ExcelAddIn.Services;
using Excel = Microsoft.Office.Interop.Excel;
using WinFormsTimer = System.Windows.Forms.Timer;

namespace LudicrousSpeed.ExcelAddIn
{
    [ComVisible(true)]
    public sealed class LudicrousSpeedRibbon : ExcelRibbon, IExcelAddIn
    {
        private readonly NativeEngineClient engineClient = new NativeEngineClient();
        private readonly WorkbookChangeTracker changeTracker = new WorkbookChangeTracker();
        private readonly WorkbookSnapshotService snapshotService;
        private readonly LiveValuePublisher livePublisher = new LiveValuePublisher();
        private readonly ReportSheetWriter reportWriter = new ReportSheetWriter();
        private readonly DataTableConverter dataTableConverter;

        /// <summary>
        /// Opt-in only: set LUDICROUS_ASYNC_RUN=1 to run the native engine
        /// call off Excel's UI thread instead of blocking it. Off by default
        /// until verified against live Excel -- see the doc comment on
        /// <see cref="RunAsync"/>.
        /// </summary>
        private static bool AsyncRunEnabled =>
            Environment.GetEnvironmentVariable("LUDICROUS_ASYNC_RUN") == "1";

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

        public LudicrousSpeedRibbon()
        {
            snapshotService = new WorkbookSnapshotService(changeTracker);
            dataTableConverter = new DataTableConverter(changeTracker);
        }

        public void AutoOpen()
        {
            changeTracker.Start();
            CalculationKeyBinding.Initialize(() => Run("recalculate", includeExcelBaseline: false));
        }

        public void AutoClose()
        {
            CalculationKeyBinding.Shutdown();
            changeTracker.Stop();
        }

        public void ToggleInterceptF9(IRibbonControl control, bool pressed)
        {
            CalculationKeyBinding.SetEnabled(pressed);
        }

        public bool GetInterceptF9Pressed(IRibbonControl control)
        {
            return CalculationKeyBinding.Enabled;
        }

        public override string GetCustomUI(string ribbonId)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='LudicrousSpeedTab' label='LudicrousSpeed'>
        <group id='LudicrousSpeedCalcGroup' label='Calculation'>
          <button id='AnalyzeWorkbookButton'
                  label='Analyze Workbook'
                  size='large'
                  imageMso='ErrorChecking'
                  onAction='AnalyzeWorkbook'
                  screentip='Analyze Workbook'
                  supertip='Scan the active workbook and report IronCalc coverage and fallback regions.' />
          <button id='RecalculateWorkbookButton'
                  label='Recalculate with LudicrousSpeed'
                  size='large'
                  imageMso='CalculateNow'
                  onAction='RecalculateWithLudicrousSpeed'
                  screentip='Recalculate with LudicrousSpeed'
                  supertip='Evaluate the active workbook with IronCalc first and Excel fallback for unsupported regions.' />
          <button id='BenchmarkWorkbookButton'
                  label='Benchmark'
                  size='large'
                  imageMso='CalculateSheet'
                  onAction='BenchmarkWorkbook'
                  screentip='Benchmark Workbook'
                  supertip='Compare Excel full rebuild timing against the LudicrousSpeed prototype engine.' />
          <toggleButton id='InterceptF9Toggle'
                  label='F9 Uses LudicrousSpeed'
                  imageMso='CalculateNow'
                  onAction='ToggleInterceptF9'
                  getPressed='GetInterceptF9Pressed'
                  screentip='Route F9 to the LudicrousSpeed engine'
                  supertip='Press F9 to recalculate with LudicrousSpeed instead of Excel. Shift+F9 and Ctrl+Alt+F9 still run Excel&apos;s own calculation. Setting is remembered between sessions.' />
          <toggleButton id='DetailedReportToggle'
                  label='Detailed Report'
                  imageMso='ErrorChecking'
                  onAction='ToggleDetailedReport'
                  getPressed='GetDetailedReportPressed'
                  screentip='Detailed Report (audit mode)'
                  supertip='Include per-fallback samples, data table diagnostics, writeback failures and the fallback detail sheet. Slower on large models, so off by default.' />
        </group>
        <group id='LudicrousSpeedDataTableGroup' label='Data Tables'>
          <button id='ConvertDataTablesButton'
                  label='Convert to Live'
                  size='large'
                  imageMso='Refresh'
                  onAction='ConvertDataTablesToLive'
                  screentip='Convert Data Tables to LudicrousSpeed Live Cells'
                  supertip='Replace native Excel data tables with LS.LIVE cells driven by the LudicrousSpeed kernel. Excel stops re-running the table once per scenario; the source formula and axis inputs are left untouched.' />
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

        public void RecalculateWithLudicrousSpeed(IRibbonControl control)
        {
            Run("recalculate", includeExcelBaseline: false);
        }

        public void BenchmarkWorkbook(IRibbonControl control)
        {
            Run("benchmark", includeExcelBaseline: true);
        }

        /// <summary>
        /// Runs the engine to discover this workbook's native data tables,
        /// then replaces each eligible one with LS.LIVE cells. The engine run
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

                // Publish first: the LS.LIVE formulas written below resolve
                // immediately instead of showing #N/A until the next recalc.
                var published = livePublisher.Publish(response);

                var regions = response.Result.Benchmark.DataTables.Regions;
                var confirm = MessageBox.Show(
                    $"Convert {regions.Count} native data table(s) to LudicrousSpeed live cells?"
                        + Environment.NewLine + Environment.NewLine
                        + "The Excel data tables will be replaced with LS.LIVE formulas. Your source "
                        + "formulas and axis inputs are not modified, and 'Restore Native' puts the "
                        + "original tables back."
                        + Environment.NewLine + Environment.NewLine
                        + $"{published:N0} engine values are available to drive them.",
                    "LudicrousSpeed",
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

                MessageBox.Show(message, "LudicrousSpeed", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                MessageBox.Show(message, "LudicrousSpeed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
        }

        /// <summary>
        /// True from the moment a run starts until it has finished reporting.
        /// Only ever read or written on Excel's UI thread -- both entry points
        /// (ribbon callbacks and the F9 macro) arrive there, and the async
        /// path's completion is marshalled back via QueueAsMacro -- so it needs
        /// no synchronisation.
        ///
        /// The sync path is self-limiting: it blocks the UI thread, so a second
        /// run cannot be started while one is going. The async path is not, and
        /// F9 makes starting one trivial -- a held-down key would otherwise
        /// stack engine runs that then serialise behind the engine's own lock.
        /// </summary>
        private bool runInFlight;

        private void Run(string mode, bool includeExcelBaseline)
        {
            if (runInFlight)
            {
                SetStatusBar("LudicrousSpeed is already calculating...");
                return;
            }

            runInFlight = true;
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
                var ludicrousStopwatch = Stopwatch.StartNew();
                snapshot = snapshotService.Create(mode, excelBaselineMs);
                snapshot.DataTableOverrides = dataTableConverter.ReadOverrides();
                snapshot.IncludeAnalytics = detailedReport;
                var response = engineClient.Run(snapshot, out var nativeCallMs);
                ludicrousStopwatch.Stop();

                FinishRun(mode, snapshot, response, excelBaselineMs, ludicrousStopwatch.ElapsedMilliseconds, nativeCallMs);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
            finally
            {
                runInFlight = false;
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
            // Declared out here so the catch below can stop it: the failure it
            // handles can happen after the timer has started.
            WinFormsTimer? progressTimer = null;
            try
            {
                var excelBaselineMs = MeasureExcelBaseline(includeExcelBaseline);
                var ludicrousStopwatch = Stopwatch.StartNew();
                snapshot = snapshotService.Create(mode, excelBaselineMs);
                snapshot.DataTableOverrides = dataTableConverter.ReadOverrides();
                snapshot.IncludeAnalytics = detailedReport;
                var snapshotForBackgroundCall = snapshot;

                SetStatusBar("LudicrousSpeed is calculating in the background...");

                // Excel writes "Calculating (8 processors): 42%" to the status
                // bar during its own recalculation, so a long run that reports
                // nothing reads as a hang. This only works on the async path:
                // RunSync holds the UI thread inside the native call, so no
                // timer can tick and the status bar cannot repaint until the
                // run is already over.
                progressTimer = StartProgressTimer();

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

                            ludicrousStopwatch.Stop();
                            FinishRun(
                                mode,
                                snapshotForBackgroundCall,
                                response!,
                                excelBaselineMs,
                                ludicrousStopwatch.ElapsedMilliseconds,
                                nativeCallMs);
                        }
                        catch (Exception ex)
                        {
                            ShowError(ex.Message);
                        }
                        finally
                        {
                            // Runs on Excel's UI thread via QueueAsMacro, which
                            // is what makes clearing the flag here safe.
                            runInFlight = false;
                            StopProgressTimer(progressTimer);
                            ClearStatusBar();
                            calculationGuard.Dispose();
                            TryDeleteSnapshot(snapshotForBackgroundCall);
                        }
                    });
                });
            }
            catch (Exception ex)
            {
                // Reached only if the failure happened before the background
                // task was handed off, so nothing else will clear this.
                runInFlight = false;
                StopProgressTimer(progressTimer);
                ClearStatusBar();
                calculationGuard.Dispose();
                TryDeleteSnapshot(snapshot);
                ShowError(ex.Message);
            }
        }

        /// <summary>
        /// Polls the engine's progress counters from Excel's UI thread. A
        /// WinForms timer rather than a background loop precisely because its
        /// tick arrives on the UI thread, which is the only place the Excel
        /// object model may be touched -- so the status bar write needs no
        /// marshalling.
        ///
        /// 250ms is fast enough to look live and slow enough that the poll
        /// itself never shows up next to a multi-second calculation.
        /// </summary>
        private WinFormsTimer StartProgressTimer()
        {
            var timer = new WinFormsTimer { Interval = 250 };
            timer.Tick += (sender, args) =>
            {
                try
                {
                    var progress = engineClient.ReadProgress();
                    if (progress.IsRunning)
                    {
                        SetStatusBar(progress.Describe());
                    }
                }
                catch
                {
                    // Never let a progress read break a run that is working.
                }
            };
            timer.Start();
            return timer;
        }

        private static void StopProgressTimer(WinFormsTimer? timer)
        {
            if (timer == null)
            {
                return;
            }

            try
            {
                timer.Stop();
                timer.Dispose();
            }
            catch
            {
                // Nothing useful to do; the timer dies with the add-in.
            }
        }

        /// <summary>
        /// Everything after the native engine call has returned a response:
        /// applying writeback, writing the report sheet, and showing the
        /// completion message. Shared by both the sync and async paths so
        /// enabling LUDICROUS_ASYNC_RUN changes only how/when this runs, not
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
            // Push engine results to any LS.LIVE cells, then let Excel
            // propagate from those injection points. RTD values land
            // asynchronously, so the recalculation is queued rather than
            // called inline -- QueueAsMacro runs it once Excel has processed
            // the pending updates, otherwise Calculate could run against
            // values that haven't arrived yet and (under Manual mode) never
            // get recalculated again.
            var livePublished = livePublisher.Publish(response);
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
                        // A failed propagation pass leaves the LS.LIVE cells
                        // themselves correct; the user can recalculate again.
                        // Safe even with F9 intercepted -- that path runs this
                        // same pipeline, and re-publishing is idempotent.
                    }
                });
            }

            var hostMetrics = new HostRunMetrics
            {
                ExcelBaselineMs = excelBaselineMs,
                SnapshotSaveMs = snapshot.SnapshotSaveMs,
                SnapshotSkipped = snapshot.SnapshotSkipped,
                NativeCallMs = nativeCallMs,
                LudicrousEndToEndMs = warpSpeedEndToEndMs,
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
            var elapsed = hostMetrics.LudicrousEndToEndMs;
            var summary = $"LudicrousSpeed: {elapsed:N0} ms";
            if (livePublished > 0)
            {
                summary += $", {livePublished:N0} live values published";
            }

            if (detailedReport)
            {
                summary += " - see _LudicrousSpeed_Report";
            }

            SetStatusBar(summary);
        }

        private static void ShowError(string? message)
        {
            MessageBox.Show(message, "LudicrousSpeed", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        /// compares LudicrousSpeed-with-tables against Excel-without-them, which
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
