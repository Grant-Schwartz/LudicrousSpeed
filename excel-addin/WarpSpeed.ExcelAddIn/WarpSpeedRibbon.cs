using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using WarpSpeed.ExcelAddIn.Interop;
using WarpSpeed.ExcelAddIn.Models;
using WarpSpeed.ExcelAddIn.Services;

namespace WarpSpeed.ExcelAddIn
{
    public sealed class WarpSpeedRibbon : ExcelRibbon, IExcelAddIn
    {
        private readonly NativeEngineClient engineClient = new NativeEngineClient();
        private readonly WorkbookChangeTracker changeTracker = new WorkbookChangeTracker();
        private readonly WorkbookSnapshotService snapshotService;
        private readonly ReportSheetWriter reportWriter = new ReportSheetWriter();

        public WarpSpeedRibbon()
        {
            snapshotService = new WorkbookSnapshotService(changeTracker);
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
                  imageMso='ReviewInspectDocument'
                  size='large'
                  onAction='AnalyzeWorkbook'
                  screentip='Analyze Workbook'
                  supertip='Scan the active workbook and report IronCalc coverage and fallback regions.' />
          <button id='RecalculateWorkbookButton'
                  label='Recalculate with WarpSpeed'
                  imageMso='CalculateNow'
                  size='large'
                  onAction='RecalculateWithWarpSpeed'
                  screentip='Recalculate with WarpSpeed'
                  supertip='Evaluate the active workbook with IronCalc first and Excel fallback for unsupported regions.' />
          <button id='BenchmarkWorkbookButton'
                  label='Benchmark'
                  imageMso='Gauge'
                  size='large'
                  onAction='BenchmarkWorkbook'
                  screentip='Benchmark Workbook'
                  supertip='Compare Excel full rebuild timing against the WarpSpeed prototype engine.' />
          <button id='RestoreButton'
                  label='Restore Last Results'
                  imageMso='Undo'
                  onAction='RestoreLastResults'
                  screentip='Restore Last Results'
                  supertip='Restore value changes made by the last writeback-capable WarpSpeed run.' />
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

        public void RestoreLastResults(IRibbonControl control)
        {
            MessageBox.Show(
                "Restore is wired into the ribbon, but value writeback is disabled in this prototype so there is nothing to restore yet.",
                "WarpSpeed",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void Run(string mode, bool includeExcelBaseline)
        {
            WorkbookSnapshot? snapshot = null;
            try
            {
                var excelBaselineMs = MeasureExcelBaseline(includeExcelBaseline);
                var warpspeedStopwatch = Stopwatch.StartNew();
                snapshot = snapshotService.Create(mode, excelBaselineMs);
                var response = engineClient.Run(snapshot, out var nativeCallMs);
                warpspeedStopwatch.Stop();

                var hostMetrics = new HostRunMetrics
                {
                    ExcelBaselineMs = excelBaselineMs,
                    SnapshotSaveMs = snapshot.SnapshotSaveMs,
                    SnapshotSkipped = snapshot.SnapshotSkipped,
                    NativeCallMs = nativeCallMs,
                    WarpSpeedEndToEndMs = warpspeedStopwatch.ElapsedMilliseconds,
                };

                reportWriter.Write(response, hostMetrics);

                if (!response.Ok)
                {
                    MessageBox.Show(response.Error, "WarpSpeed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                changeTracker.MarkRunSucceeded(snapshot);

                MessageBox.Show(
                    "WarpSpeed completed. See the _WarpSpeed_Report sheet for coverage, fallback, and timing details.",
                    "WarpSpeed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "WarpSpeed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                TryDeleteSnapshot(snapshot);
            }
        }

        private static long? MeasureExcelBaseline(bool includeExcelBaseline)
        {
            if (!includeExcelBaseline)
            {
                return null;
            }

            dynamic excel = ExcelDnaUtil.Application;
            var stopwatch = Stopwatch.StartNew();
            excel.CalculateFullRebuild();
            stopwatch.Stop();
            return stopwatch.ElapsedMilliseconds;
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
