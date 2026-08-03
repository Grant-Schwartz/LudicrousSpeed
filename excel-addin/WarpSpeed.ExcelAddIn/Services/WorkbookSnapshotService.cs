using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using WarpSpeed.ExcelAddIn.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace WarpSpeed.ExcelAddIn.Services
{
    internal sealed class WorkbookSnapshotService
    {
        private readonly WorkbookChangeTracker changeTracker;

        public WorkbookSnapshotService(WorkbookChangeTracker changeTracker)
        {
            this.changeTracker = changeTracker;
        }

        public WorkbookSnapshot Create(string mode, long? excelBaselineMs)
        {
            dynamic excel = ExcelDnaUtil.Application;
            Excel.Workbook workbook = excel.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("Open a workbook before running WarpSpeed.");
            }

            var changeSet = changeTracker.CaptureForSnapshot((object)workbook);
            var changedCells = new List<ChangedCell>();
            var forceReload = changeSet.ForceReload;

            if (!forceReload && changeSet.IsWarm)
            {
                try
                {
                    changedCells = MaterializeChangedCells((Excel.Workbook)workbook, changeSet.DirtyCells);
                }
                catch
                {
                    forceReload = true;
                    changedCells.Clear();
                }
            }

            var needsSnapshot = forceReload || !changeSet.IsWarm;
            var tempPath = "";
            var snapshotSaveMs = 0L;
            if (needsSnapshot)
            {
                tempPath = Path.Combine(
                    Path.GetTempPath(),
                    $"warpspeed-{Guid.NewGuid():N}.xlsx");

                var snapshotStopwatch = Stopwatch.StartNew();
                SaveWorkbookCopy(workbook, tempPath);
                snapshotStopwatch.Stop();
                snapshotSaveMs = snapshotStopwatch.ElapsedMilliseconds;
            }

            return new WorkbookSnapshot
            {
                WorkbookPath = tempPath,
                WorkbookName = Convert.ToString(workbook.Name),
                WorkbookId = changeSet.WorkbookId,
                Mode = mode,
                ExcelBaselineMs = excelBaselineMs,
                ForceReload = forceReload,
                ChangedCells = changedCells,
                EvaluateDataTables = false,
                Locale = "en",
                Timezone = "UTC",
                Language = "en",
                SnapshotSaveMs = snapshotSaveMs,
                SnapshotSkipped = !needsSnapshot,
                SheetSignature = changeSet.SheetSignature,
            };
        }

        private static List<ChangedCell> MaterializeChangedCells(
            Excel.Workbook workbook,
            IEnumerable<DirtyCellAddress> dirtyCells)
        {
            var changedCells = new List<ChangedCell>();
            foreach (var dirtyCell in dirtyCells)
            {
                var worksheet = FindWorksheet(workbook, dirtyCell.SheetName);
                dynamic cell = worksheet.Cells[dirtyCell.Row, dirtyCell.Column];
                var hasFormula = Convert.ToBoolean(cell.HasFormula, CultureInfo.InvariantCulture);
                var input = hasFormula
                    ? Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? ""
                    : ConvertCellValue(cell.Value2);

                changedCells.Add(new ChangedCell
                {
                    SheetName = dirtyCell.SheetName,
                    Row = dirtyCell.Row,
                    Column = dirtyCell.Column,
                    Address = dirtyCell.Address,
                    Input = input,
                    IsFormula = hasFormula,
                });
            }

            return changedCells;
        }

        private static Excel.Worksheet FindWorksheet(Excel.Workbook workbook, string sheetName)
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

            throw new InvalidOperationException($"Worksheet not found: {sheetName}");
        }

        private static string ConvertCellValue(object value)
        {
            if (value == null)
            {
                return "";
            }

            if (value is ErrorWrapper)
            {
                throw new InvalidOperationException("Excel error values require a full workbook snapshot.");
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        }

        private static void SaveWorkbookCopy(Excel.Workbook workbook, string tempPath)
        {
            try
            {
                workbook.SaveCopyAs(tempPath);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    "Excel could not save a temporary workbook copy for WarpSpeed. Save the workbook as an .xlsx file, close Protected View or permission prompts, and try again. Details: " + ex.Message,
                    ex);
            }
        }
    }
}
