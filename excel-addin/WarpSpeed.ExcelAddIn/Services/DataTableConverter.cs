using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;
using WarpSpeed.ExcelAddIn.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace WarpSpeed.ExcelAddIn.Services
{
    /// <summary>
    /// Replaces native Excel data tables with WS.LIVE cells driven by the
    /// WarpSpeed kernel, and puts them back on request.
    ///
    /// WHY: a data table is the most expensive structure in a model. Excel
    /// re-evaluates the table's source formula cone once per scenario cell,
    /// so a 5x5 two-variable table costs 25 full passes. Removing the native
    /// table removes that cost outright; the kernel computes the same grid in
    /// one parallel pass and pushes the values over RTD.
    ///
    /// WHY THIS IS LESS INVASIVE THAN IT SOUNDS: the cells being replaced hold
    /// Excel-generated {=TABLE(r1,r2)} array markers, not anything a modeller
    /// wrote. The source formula cell and the row/column axis inputs -- the
    /// parts someone actually authored -- are never touched.
    ///
    /// An Excel data table is one array formula over the whole body range, so
    /// individual cells can't be swapped; the array is cleared and per-cell
    /// WS.LIVE formulas are written in its place. Restore metadata is kept on
    /// a hidden worksheet rather than in memory, so a restore still works in a
    /// later session or after a crash.
    ///
    /// NOT YET VERIFIED AGAINST LIVE EXCEL.
    /// </summary>
    internal sealed class DataTableConverter
    {
        private const string MetadataSheetName = WorkbookChangeTracker.DataTableSheetName;
        private const string LiveFunction = "WS.LIVE";

        private readonly WorkbookChangeTracker changeTracker;

        public DataTableConverter(WorkbookChangeTracker changeTracker)
        {
            this.changeTracker = changeTracker;
        }

        public ConversionResult ConvertToLive(IReadOnlyList<DataTableRegionInfo> regions)
        {
            var result = new ConversionResult();
            if (regions == null || regions.Count == 0)
            {
                result.Message = "The engine reported no native data tables in this workbook.";
                return result;
            }

            dynamic excel = ExcelDnaUtil.Application;
            Excel.Workbook workbook = excel.ActiveWorkbook;
            var previousCalculation = excel.Calculation;
            var previousScreenUpdating = excel.ScreenUpdating;

            try
            {
                // Manual mode for the duration: clearing an array formula and
                // writing hundreds of cells would otherwise trigger a cascade
                // of recalculations, which is exactly the cost being removed.
                excel.Calculation = Excel.XlCalculation.xlCalculationManual;
                excel.ScreenUpdating = false;

                // Conversion writes over a thousand formulas plus the metadata
                // sheet. None of that is a model edit, and letting it reach the
                // change tracker both floods the dirty-cell budget (forcing a
                // needless full reload next run) and records edits on a sheet
                // no cached engine model knows about.
                using var trackingSuspension = changeTracker.SuspendTracking();

                var metadata = EnsureMetadataSheet(workbook);

                foreach (var region in regions)
                {
                    if (!region.KernelEligible)
                    {
                        // No kernel values for this shape, so its native table
                        // has to stay -- converting it would leave dead cells.
                        result.Skipped++;
                        result.SkippedReasons.Add(
                            $"{region.TableId}: kernel cannot evaluate this table's shape");
                        continue;
                    }

                    try
                    {
                        ConvertOne(workbook, metadata, region);
                        result.Converted++;
                        result.ConvertedCells += region.CellCount;
                    }
                    catch (Exception ex)
                    {
                        result.Failed++;
                        result.SkippedReasons.Add($"{region.TableId}: {ex.Message}");
                    }
                }

                result.Message =
                    $"Converted {result.Converted} data table(s) ({result.ConvertedCells:N0} cells) to "
                    + $"WarpSpeed live cells. Skipped {result.Skipped}, failed {result.Failed}.";
            }
            finally
            {
                try { excel.ScreenUpdating = previousScreenUpdating; } catch { }
                try { excel.Calculation = previousCalculation; } catch { }
            }

            return result;
        }

        private static void ConvertOne(
            Excel.Workbook workbook,
            Excel.Worksheet metadata,
            DataTableRegionInfo region)
        {
            Excel.Worksheet sheet = FindWorksheet(workbook, region.SheetName)
                ?? throw new InvalidOperationException($"sheet '{region.SheetName}' not found");

            Excel.Range body = sheet.Range[region.RangeAddress];

            // Record before destroying anything, so a restore is possible even
            // if the write below fails partway.
            RecordRestoreInfo(metadata, region);

            // Clearing the range removes the {=TABLE()} array formula. Excel
            // refuses to modify part of an array, so the whole body goes at
            // once.
            body.ClearContents();

            var firstRow = body.Row;
            var firstColumn = body.Column;
            for (var r = 0; r < body.Rows.Count; r++)
            {
                for (var c = 0; c < body.Columns.Count; c++)
                {
                    Excel.Range cell = (Excel.Range)sheet.Cells[firstRow + r, firstColumn + c];
                    var address = ColumnLetters(firstColumn + c)
                        + (firstRow + r).ToString(CultureInfo.InvariantCulture);
                    // Sheet name is quoted because model sheets routinely
                    // contain spaces and parentheses, e.g. LBO (Share Price).
                    cell.Formula =
                        $"={LiveFunction}(\"{EscapeForFormula(region.SheetName)}!{address}\")";
                }
            }
        }

        /// <summary>
        /// Reads back every table this workbook has had converted, so the
        /// definitions can ride along on each snapshot. Without this the
        /// engine finds no table where a converted one used to be, computes
        /// nothing for it, and its live cells never update again.
        /// </summary>
        public List<DataTableOverride> ReadOverrides()
        {
            var overrides = new List<DataTableOverride>();
            try
            {
                dynamic excel = ExcelDnaUtil.Application;
                Excel.Workbook workbook = excel.ActiveWorkbook;
                Excel.Worksheet? metadata = FindWorksheet(workbook, MetadataSheetName);
                if (metadata == null)
                {
                    return overrides;
                }

                var row = 2; // row 1 is the header
                while (true)
                {
                    var tableId = CellText(metadata, row, 1);
                    if (string.IsNullOrWhiteSpace(tableId))
                    {
                        break;
                    }

                    var sheetName = CellText(metadata, row, 2);
                    var rangeAddress = CellText(metadata, row, 3);
                    if (!string.IsNullOrWhiteSpace(sheetName)
                        && !string.IsNullOrWhiteSpace(rangeAddress))
                    {
                        var columnInput = CellText(metadata, row, 4);
                        var rowInput = CellText(metadata, row, 5);
                        overrides.Add(new DataTableOverride
                        {
                            SheetName = sheetName,
                            RangeAddress = rangeAddress,
                            AnchorAddress = CellText(metadata, row, 8),
                            ColumnInputCell = string.IsNullOrWhiteSpace(columnInput) ? null : columnInput,
                            RowInputCell = string.IsNullOrWhiteSpace(rowInput) ? null : rowInput,
                            IsTwoDimensional = CellText(metadata, row, 7) == "1",
                            // Excel only exposes dtr for one-variable tables;
                            // a two-variable table always uses both axes.
                            Dtr = CellText(metadata, row, 7) == "1",
                        });
                    }

                    row++;
                }
            }
            catch (Exception)
            {
                // Overrides are an optimization for already-converted tables;
                // a failure to read them must not take down the whole run.
            }

            return overrides;
        }

        /// <summary>
        /// Forgets a restored table so its definition stops being replayed.
        /// </summary>
        private static void ClearMetadataRow(Excel.Worksheet metadata, int row)
        {
            for (var column = 1; column <= 8; column++)
            {
                ((Excel.Range)metadata.Cells[row, column]).ClearContents();
            }
        }

        public RestoreResult RestoreNativeTables()
        {
            var result = new RestoreResult();
            dynamic excel = ExcelDnaUtil.Application;
            Excel.Workbook workbook = excel.ActiveWorkbook;
            Excel.Worksheet? metadata = FindWorksheet(workbook, MetadataSheetName);
            if (metadata == null)
            {
                result.Message = "No WarpSpeed data table conversions are recorded in this workbook.";
                return result;
            }

            var previousCalculation = excel.Calculation;
            var previousScreenUpdating = excel.ScreenUpdating;
            try
            {
                excel.Calculation = Excel.XlCalculation.xlCalculationManual;
                excel.ScreenUpdating = false;
                using var trackingSuspension = changeTracker.SuspendTracking();

                var row = 2; // row 1 is the header
                while (true)
                {
                    var tableId = Convert.ToString(
                        ((Excel.Range)metadata.Cells[row, 1]).Value2,
                        CultureInfo.InvariantCulture) ?? "";
                    if (string.IsNullOrWhiteSpace(tableId))
                    {
                        break;
                    }

                    try
                    {
                        RestoreOne(workbook, metadata, row);
                        // The native table is back, so the engine will
                        // discover it again; replaying the override would now
                        // shadow the real thing.
                        ClearMetadataRow(metadata, row);
                        result.Restored++;
                    }
                    catch (Exception ex)
                    {
                        result.Failed++;
                        result.Errors.Add($"{tableId}: {ex.Message}");
                    }

                    row++;
                }

                result.Message = $"Restored {result.Restored} native data table(s). Failed {result.Failed}.";
            }
            finally
            {
                try { excel.ScreenUpdating = previousScreenUpdating; } catch { }
                try { excel.Calculation = previousCalculation; } catch { }
            }

            return result;
        }

        private static void RestoreOne(Excel.Workbook workbook, Excel.Worksheet metadata, int row)
        {
            var sheetName = CellText(metadata, row, 2);
            var rangeAddress = CellText(metadata, row, 3);
            var columnInput = CellText(metadata, row, 4);
            var rowInput = CellText(metadata, row, 5);

            Excel.Worksheet sheet = FindWorksheet(workbook, sheetName)
                ?? throw new InvalidOperationException($"sheet '{sheetName}' not found");

            Excel.Range body = sheet.Range[rangeAddress];
            body.ClearContents();

            // Excel's Data Table dialog takes (RowInput, ColumnInput); a
            // one-variable table passes only the one it uses, so the unused
            // side is omitted rather than passed as an empty reference.
            var rowInputRange = ResolveCell(workbook, rowInput);
            var columnInputRange = ResolveCell(workbook, columnInput);

            if (rowInputRange != null && columnInputRange != null)
            {
                body.Table(rowInputRange, columnInputRange);
            }
            else if (columnInputRange != null)
            {
                body.Table(ColumnInput: columnInputRange);
            }
            else if (rowInputRange != null)
            {
                body.Table(RowInput: rowInputRange);
            }
            else
            {
                throw new InvalidOperationException("no input cell recorded for this table");
            }
        }

        private static void RecordRestoreInfo(Excel.Worksheet metadata, DataTableRegionInfo region)
        {
            var row = NextMetadataRow(metadata);
            ((Excel.Range)metadata.Cells[row, 1]).Value2 = region.TableId;
            ((Excel.Range)metadata.Cells[row, 2]).Value2 = region.SheetName;
            ((Excel.Range)metadata.Cells[row, 3]).Value2 = region.RangeAddress;
            ((Excel.Range)metadata.Cells[row, 4]).Value2 = region.ColumnInputCell ?? "";
            ((Excel.Range)metadata.Cells[row, 5]).Value2 = region.RowInputCell ?? "";
            ((Excel.Range)metadata.Cells[row, 6]).Value2 =
                DateTime.Now.ToString("u", CultureInfo.InvariantCulture);
            // dt2D/dtr are needed to rebuild the region shape engine-side on
            // every later run, not just to restore the native table.
            ((Excel.Range)metadata.Cells[row, 7]).Value2 = region.IsTwoDimensional ? 1d : 0d;
            ((Excel.Range)metadata.Cells[row, 8]).Value2 = AnchorOf(region.RangeAddress);
        }

        private static int NextMetadataRow(Excel.Worksheet metadata)
        {
            var row = 2;
            while (!string.IsNullOrWhiteSpace(CellText(metadata, row, 1)))
            {
                row++;
            }

            return row;
        }

        private static Excel.Worksheet EnsureMetadataSheet(Excel.Workbook workbook)
        {
            Excel.Worksheet? existing = FindWorksheet(workbook, MetadataSheetName);
            if (existing != null)
            {
                return existing;
            }

            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.Add();
            sheet.Name = MetadataSheetName;
            ((Excel.Range)sheet.Cells[1, 1]).Value2 = "table_id";
            ((Excel.Range)sheet.Cells[1, 2]).Value2 = "sheet_name";
            ((Excel.Range)sheet.Cells[1, 3]).Value2 = "range_address";
            ((Excel.Range)sheet.Cells[1, 4]).Value2 = "column_input_cell";
            ((Excel.Range)sheet.Cells[1, 5]).Value2 = "row_input_cell";
            ((Excel.Range)sheet.Cells[1, 6]).Value2 = "converted_at";
            ((Excel.Range)sheet.Cells[1, 7]).Value2 = "is_two_dimensional";
            ((Excel.Range)sheet.Cells[1, 8]).Value2 = "anchor_address";
            sheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            return sheet;
        }

        private static Excel.Range? ResolveCell(Excel.Workbook workbook, string qualifiedAddress)
        {
            if (string.IsNullOrWhiteSpace(qualifiedAddress))
            {
                return null;
            }

            var bang = qualifiedAddress.LastIndexOf('!');
            if (bang <= 0)
            {
                return null;
            }

            var sheetName = qualifiedAddress.Substring(0, bang).Trim().Trim('\'');
            var cell = qualifiedAddress.Substring(bang + 1).Trim();
            Excel.Worksheet? sheet = FindWorksheet(workbook, sheetName);
            return sheet?.Range[cell];
        }

        private static string CellText(Excel.Worksheet sheet, int row, int column)
        {
            return Convert.ToString(
                ((Excel.Range)sheet.Cells[row, column]).Value2,
                CultureInfo.InvariantCulture) ?? "";
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

        /// <summary>
        /// Top-left cell of a body range -- where Excel keeps the dataTable
        /// formula. Derived rather than carried on DataTableRegionInfo, since
        /// it is always just the first half of the range.
        /// </summary>
        private static string AnchorOf(string rangeAddress)
        {
            var colon = rangeAddress.IndexOf(':');
            return colon < 0 ? rangeAddress : rangeAddress.Substring(0, colon);
        }

        private static string EscapeForFormula(string sheetName)
        {
            return sheetName.Replace("\"", "\"\"");
        }

        private static string ColumnLetters(int oneBasedColumn)
        {
            var letters = "";
            var remaining = oneBasedColumn;
            while (remaining > 0)
            {
                var offset = (remaining - 1) % 26;
                letters = (char)('A' + offset) + letters;
                remaining = (remaining - 1) / 26;
            }

            return letters;
        }
    }

    internal sealed class ConversionResult
    {
        public int Converted { get; set; }
        public int ConvertedCells { get; set; }
        public int Skipped { get; set; }
        public int Failed { get; set; }
        public List<string> SkippedReasons { get; } = new List<string>();
        public string Message { get; set; } = "";
    }

    internal sealed class RestoreResult
    {
        public int Restored { get; set; }
        public int Failed { get; set; }
        public List<string> Errors { get; } = new List<string>();
        public string Message { get; set; } = "";
    }
}
