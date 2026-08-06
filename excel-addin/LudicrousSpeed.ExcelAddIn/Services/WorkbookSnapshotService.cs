using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using LudicrousSpeed.ExcelAddIn.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace LudicrousSpeed.ExcelAddIn.Services
{
    internal sealed class WorkbookSnapshotService
    {
        private readonly WorkbookChangeTracker changeTracker;

        public WorkbookSnapshotService(WorkbookChangeTracker changeTracker)
        {
            this.changeTracker = changeTracker;
        }

        /// <summary>
        /// Opt-in only: set LUDICROUS_INLINE_SNAPSHOT=1 to try the in-memory
        /// (no SaveCopyAs, no xlsx re-import) cold-load path instead of the
        /// file-based one. Off by default until InMemoryWorkbookReader has
        /// been verified against live Excel -- see its doc comment.
        /// </summary>
        private static bool InlineSnapshotEnabled =>
            Environment.GetEnvironmentVariable("LUDICROUS_INLINE_SNAPSHOT") == "1";

        public WorkbookSnapshot Create(string mode, long? excelBaselineMs)
        {
            dynamic excel = ExcelDnaUtil.Application;
            Excel.Workbook workbook = excel.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("Open a workbook before running LudicrousSpeed.");
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
            InlineWorkbook? inlineWorkbook = null;
            if (needsSnapshot)
            {
                var snapshotStopwatch = Stopwatch.StartNew();
                if (InlineSnapshotEnabled)
                {
                    try
                    {
                        inlineWorkbook = InMemoryWorkbookReader.Read(workbook);
                    }
                    catch
                    {
                        // Fall back to the proven file-based path below for
                        // any read failure (unsupported cell shape, COM
                        // error, etc.) rather than failing the whole run.
                        inlineWorkbook = null;
                    }
                }

                if (inlineWorkbook == null)
                {
                    tempPath = Path.Combine(
                        Path.GetTempPath(),
                        $"ludicrous-{Guid.NewGuid():N}.xlsx");
                    SaveWorkbookCopy(workbook, tempPath);
                }
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
                EvaluateDataTables = true,
                Locale = "en",
                Timezone = "UTC",
                Language = "en",
                InlineWorkbook = inlineWorkbook,
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
                    "Excel could not save a temporary workbook copy for LudicrousSpeed. Save the workbook as an .xlsx file, close Protected View or permission prompts, and try again. Details: " + ex.Message,
                    ex);
            }
        }
    }
}
