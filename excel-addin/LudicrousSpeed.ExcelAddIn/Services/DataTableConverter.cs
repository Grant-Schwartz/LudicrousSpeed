using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;
using LudicrousSpeed.ExcelAddIn.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace LudicrousSpeed.ExcelAddIn.Services
{
    /// <summary>
    /// Replaces native Excel data tables with LS.LIVE cells driven by the
    /// LudicrousSpeed kernel, and puts them back on request.
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
    /// LS.LIVE formulas are written in its place. Restore metadata is kept on
    /// a hidden worksheet rather than in memory, so a restore still works in a
    /// later session or after a crash.
    ///
    /// NOT YET VERIFIED AGAINST LIVE EXCEL.
    /// </summary>
    internal sealed class DataTableConverter
    {
        private const string MetadataSheetName = WorkbookChangeTracker.DataTableSheetName;
        private const string LiveFunction = "LS.LIVE";

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
                    + $"LudicrousSpeed live cells. Skipped {result.Skipped}, failed {result.Failed}.";
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
            RememberBody(workbook, region.TableId, body);
            // The inputs drift too, and drifted inputs are worse than a drifted
            // body: the grid still looks plausible, it is just computed from
            // the wrong assumptions.
            RememberCell(workbook, region.TableId + "|ci", region.ColumnInputCell);
            RememberCell(workbook, region.TableId + "|ri", region.RowInputCell);

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
                    // No address in the formula. Writing one as text made the
                    // cell request a fixed address forever: Excel adjusts
                    // references when rows are inserted, but not text, so the
                    // table quietly displayed correct values in the wrong
                    // places. LS.LIVE with no argument asks Excel where it is.
                    cell.Formula = $"={LiveFunction}()";
                }
            }
        }

        /// <summary>
        /// Builds an LS-native data table from a selection, the way Excel's own
        /// Data Table dialog (Alt+D+T) builds a native one -- except no native
        /// table is created at any point, so Excel never pays to re-evaluate the
        /// source formula's dependency cone once per scenario cell.
        ///
        /// The selection is the whole rectangle, exactly as Excel wants it: the
        /// source formula in the top-left corner, the column scenarios along the
        /// top row, the row scenarios down the left column. The body is
        /// therefore the selection minus its first row and first column, which
        /// is both the range the engine computes and the range that receives
        /// LS.LIVE formulas.
        ///
        /// The values do not appear until the engine has run -- the caller is
        /// expected to kick off a recalculation immediately afterwards.
        /// </summary>
        public CreateResult CreateLiveTable(
            Excel.Range selection,
            Excel.Range excelRowInput,
            Excel.Range excelColumnInput)
        {
            var result = new CreateResult();

            if (selection.Rows.Count < 2 || selection.Columns.Count < 2)
            {
                result.Message =
                    "Select the whole table first: the formula in the top-left corner, the column "
                    + "inputs along the top row, and the row inputs down the left column. That is the "
                    + "same selection Excel's own Data Table dialog expects.";
                return result;
            }

            dynamic excelApp = ExcelDnaUtil.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet sheet = (Excel.Worksheet)selection.Worksheet;
            var sheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? "";

            var firstRow = selection.Row;
            var firstColumn = selection.Column;
            var lastRow = firstRow + selection.Rows.Count - 1;
            var lastColumn = firstColumn + selection.Columns.Count - 1;

            // Without a formula in the corner there is nothing to vary, and the
            // engine would compute an empty grid. Excel rejects this case too,
            // just with a less specific message.
            var cornerFormula = Convert.ToString(
                ((Excel.Range)sheet.Cells[firstRow, firstColumn]).Formula,
                CultureInfo.InvariantCulture) ?? "";
            if (!cornerFormula.StartsWith("=", StringComparison.Ordinal))
            {
                result.Message =
                    $"{ColumnLetters(firstColumn)}{firstRow} needs to hold the formula the table varies. "
                    + (string.IsNullOrWhiteSpace(cornerFormula)
                        ? "It is currently empty."
                        : $"It currently contains '{cornerFormula}'.");
                return result;
            }

            var bodyAnchor = $"${ColumnLetters(firstColumn + 1)}${firstRow + 1}";
            var bodyAddress = $"{bodyAnchor}:${ColumnLetters(lastColumn)}${lastRow}";

            Excel.Worksheet metadata = EnsureMetadataSheet(workbook);

            // Writing LS.LIVE formulas is an edit like any other as far as the
            // tracker is concerned; without this the next run would treat the
            // whole body as user-changed input.
            using (var trackingSuspension = changeTracker.SuspendTracking())
            {
                var liveTableId = $"live-{sheetName}!{bodyAddress}";
                var existingRow = FindActiveMetadataRow(metadata, liveTableId);
                var row = existingRow > 0 ? existingRow : NextMetadataRow(metadata);
                ((Excel.Range)metadata.Cells[row, 1]).Value2 = liveTableId;
                ((Excel.Range)metadata.Cells[row, 2]).Value2 = sheetName;
                ((Excel.Range)metadata.Cells[row, 3]).Value2 = bodyAddress;
                // Excel's dialog labels and the stored fields are transposed:
                // what Excel calls the Row input cell is r1 in the OOXML, which
                // this sheet carries as column_input_cell, and vice versa. The
                // swap happens here, once, rather than in every reader -- the
                // same transposition that previously destroyed tables on
                // restore by passing the two the wrong way round.
                ((Excel.Range)metadata.Cells[row, 4]).Value2 = QualifiedAddress(excelRowInput);
                ((Excel.Range)metadata.Cells[row, 5]).Value2 = QualifiedAddress(excelColumnInput);
                ((Excel.Range)metadata.Cells[row, 6]).Value2 =
                    DateTime.Now.ToString("u", CultureInfo.InvariantCulture);
                ((Excel.Range)metadata.Cells[row, 7]).Value2 = 1d;
                ((Excel.Range)metadata.Cells[row, 8]).Value2 = bodyAnchor;
                ((Excel.Range)metadata.Cells[row, 9]).Value2 = 0d;
                // restored_at stays empty: a non-empty value there means "stop
                // replaying this as an override", which would leave the table
                // with no engine values at all.
                ((Excel.Range)metadata.Cells[row, 10]).Value2 = "";
                // Separate column, because this means something different --
                // there was never a native table here, so Restore Native has
                // nothing to rebuild and must skip it rather than fabricate one.
                ((Excel.Range)metadata.Cells[row, 11]).Value2 = "1";

                RememberBody(workbook, liveTableId, sheet.Range[bodyAddress]);
                RememberCell(workbook, liveTableId + "|ci", QualifiedAddress(excelRowInput));
                RememberCell(workbook, liveTableId + "|ri", QualifiedAddress(excelColumnInput));

                for (var r = firstRow + 1; r <= lastRow; r++)
                {
                    for (var c = firstColumn + 1; c <= lastColumn; c++)
                    {
                        ((Excel.Range)sheet.Cells[r, c]).Formula = $"={LiveFunction}()";
                    }
                }
            }

            result.Created = true;
            result.Rows = lastRow - firstRow;
            result.Columns = lastColumn - firstColumn;
            result.Message =
                $"Created a live {result.Rows} x {result.Columns} data table at {sheetName}!{bodyAddress}.";
            return result;
        }

        /// <summary>
        /// Records a table's body range as a hidden defined name, and reads it
        /// back at its current location.
        ///
        /// The stored range_address is text, so it says where the table was
        /// when it was converted, not where it is now. Excel maintains defined
        /// names across row and column insertion; text does not move. Without
        /// this the engine keeps computing the region the table used to occupy.
        /// The stored text stays as a fallback for workbooks converted before
        /// this existed, and for the case where someone deletes the name.
        /// </summary>
        private static string RangeNameFor(string tableId)
        {
            var buffer = new char[tableId.Length];
            for (var i = 0; i < tableId.Length; i++)
            {
                buffer[i] = char.IsLetterOrDigit(tableId[i]) ? tableId[i] : '_';
            }

            return "_LS_dt_" + new string(buffer);
        }

        private static void RememberBody(Excel.Workbook workbook, string tableId, Excel.Range body)
        {
            try
            {
                Excel.Name added = workbook.Names.Add(Name: RangeNameFor(tableId), RefersTo: body);
                // Hidden so it never shows up in the user's Name Manager.
                added.Visible = false;
            }
            catch (Exception)
            {
                // Falls back to the stored text, which is what happened before
                // this existed. Not worth failing a conversion over.
            }
        }

        private static void RememberCell(
            Excel.Workbook workbook,
            string key,
            string? qualifiedAddress)
        {
            if (string.IsNullOrWhiteSpace(qualifiedAddress))
            {
                return;
            }

            Excel.Range? cell = ResolveCell(workbook, qualifiedAddress!);
            if (cell != null)
            {
                RememberBody(workbook, key, cell);
            }
        }

        private static string? RememberedCellAddress(Excel.Workbook workbook, string key)
        {
            try
            {
                var target = RangeNameFor(key);
                foreach (Excel.Name candidate in workbook.Names)
                {
                    if (string.Equals(candidate.Name, target, StringComparison.OrdinalIgnoreCase))
                    {
                        return QualifiedAddress(candidate.RefersToRange);
                    }
                }
            }
            catch (Exception)
            {
                // Name pointing at deleted cells; fall back to the text.
            }

            return null;
        }

        private static string? RememberedBodyAddress(Excel.Workbook workbook, string tableId)
        {
            try
            {
                var target = RangeNameFor(tableId);
                foreach (Excel.Name candidate in workbook.Names)
                {
                    if (!string.Equals(candidate.Name, target, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    Excel.Range range = candidate.RefersToRange;
                    return "$" + ColumnLetters(range.Column) + "$" + range.Row
                        + ":$" + ColumnLetters(range.Column + range.Columns.Count - 1)
                        + "$" + (range.Row + range.Rows.Count - 1);
                }
            }
            catch (Exception)
            {
                // A name pointing at deleted cells throws on RefersToRange.
            }

            return null;
        }

        private static string QualifiedAddress(Excel.Range cell)
        {
            var sheet = (Excel.Worksheet)cell.Worksheet;
            var name = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? "";
            return $"{name}!${ColumnLetters(cell.Column)}${cell.Row}";
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
            var seenTableIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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

                    // First row wins. A workbook converted before the
                    // dedupe above can still hold several rows for one table,
                    // the older ones naming ranges that have since moved, and
                    // handing the engine all of them is what produced tables
                    // computed at their previous location.
                    if (!seenTableIds.Add(tableId))
                    {
                        row++;
                        continue;
                    }

                    var sheetName = CellText(metadata, row, 2);
                    // Current location first; the recorded text is only a
                    // fallback, because it cannot follow inserted rows.
                    var rangeAddress = RememberedBodyAddress(workbook, tableId)
                        ?? CellText(metadata, row, 3);
                    var alreadyRestored = !string.IsNullOrWhiteSpace(CellText(metadata, row, 10));
                    if (!alreadyRestored
                        && !string.IsNullOrWhiteSpace(sheetName)
                        && !string.IsNullOrWhiteSpace(rangeAddress))
                    {
                        var columnInput = RememberedCellAddress(workbook, tableId + "|ci")
                            ?? CellText(metadata, row, 4);
                        var rowInput = RememberedCellAddress(workbook, tableId + "|ri")
                            ?? CellText(metadata, row, 5);
                        overrides.Add(new DataTableOverride
                        {
                            SheetName = sheetName,
                            RangeAddress = rangeAddress,
                            // Derived, never read from column 8. The anchor is
                            // by definition the body's first cell, so reading it
                            // as separate text let it drift out of step with the
                            // range -- a fresh range against a stale anchor put
                            // the engine's idea of the source formula one row
                            // off, which is the boundary error this fixes.
                            AnchorAddress = AnchorOf(rangeAddress),
                            ColumnInputCell = string.IsNullOrWhiteSpace(columnInput) ? null : columnInput,
                            RowInputCell = string.IsNullOrWhiteSpace(rowInput) ? null : rowInput,
                            IsTwoDimensional = CellText(metadata, row, 7) == "1",
                            Dtr = CellText(metadata, row, 9) == "1",
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
        /// Marks a table as restored so its definition stops being replayed as
        /// an override, without discarding it. Deleting the row instead makes
        /// a botched restore unrecoverable -- the definition needed to rebuild
        /// the table is exactly what gets thrown away.
        /// </summary>
        private static void MarkMetadataRowRestored(Excel.Worksheet metadata, int row)
        {
            ((Excel.Range)metadata.Cells[row, 10]).Value2 =
                DateTime.Now.ToString("u", CultureInfo.InvariantCulture);
        }

        public RestoreResult RestoreNativeTables()
        {
            var result = new RestoreResult();
            dynamic excel = ExcelDnaUtil.Application;
            Excel.Workbook workbook = excel.ActiveWorkbook;
            Excel.Worksheet? metadata = FindWorksheet(workbook, MetadataSheetName);
            if (metadata == null)
            {
                result.Message = "No LudicrousSpeed data table conversions are recorded in this workbook.";
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

                    if (!string.IsNullOrWhiteSpace(CellText(metadata, row, 10)))
                    {
                        row++;
                        continue;
                    }

                    // Live-built tables are rebuilt too. Skipping them was
                    // wrong: "there was never a native table here" is true and
                    // irrelevant, because RestoreOne does not put a saved table
                    // back, it constructs one from the body range, the two
                    // input cells and the source formula -- all of which a live
                    // table records. Being able to hand the workbook to someone
                    // without the add-in is the point of the command.

                    try
                    {
                        RestoreOne(workbook, metadata, row);
                        // The native table is back, so the engine will
                        // discover it again; replaying the override would now
                        // shadow the real thing.
                        MarkMetadataRowRestored(metadata, row);
                        result.Restored++;
                    }
                    catch (Exception ex)
                    {
                        result.Failed++;
                        result.Errors.Add($"{tableId}: {ex.Message}");
                    }

                    row++;
                }

                result.Message =
                    $"Rebuilt {result.Restored} native Excel data table(s). Failed {result.Failed}.";
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
            var tableId = CellText(metadata, row, 1);
            var sheetName = CellText(metadata, row, 2);
            var rangeAddress = RememberedBodyAddress(workbook, tableId)
                ?? CellText(metadata, row, 3);
            var columnInput = RememberedCellAddress(workbook, tableId + "|ci")
                ?? CellText(metadata, row, 4);
            var rowInput = RememberedCellAddress(workbook, tableId + "|ri")
                ?? CellText(metadata, row, 5);

            Excel.Worksheet sheet = FindWorksheet(workbook, sheetName)
                ?? throw new InvalidOperationException($"sheet '{sheetName}' not found");

            var isTwoDimensional = CellText(metadata, row, 7) == "1";
            var dtr = CellText(metadata, row, 9) == "1";

            Excel.Range body = sheet.Range[rangeAddress];
            if (body.Row <= 1 || body.Column <= 1)
            {
                throw new InvalidOperationException(
                    "table body has no room for its axis row/column above and to the left");
            }

            // Only the body is cleared. The axis values and the source formula
            // sit outside it and must survive -- they are the user's own
            // content, and Table() needs them in place to rebuild from.
            body.ClearContents();

            // Range.Table() must be called on the WHOLE table rectangle --
            // the axis row above, the axis column to the left, and the body --
            // not the body alone. Calling it on just the body makes Excel
            // treat the body's own first row and column as the axes, which
            // silently produces a table one row and one column smaller filled
            // with zeros.
            Excel.Range full = sheet.Range[
                (Excel.Range)sheet.Cells[body.Row - 1, body.Column - 1],
                (Excel.Range)sheet.Cells[
                    body.Row + body.Rows.Count - 1,
                    body.Column + body.Columns.Count - 1]];

            // Our field names are inverted relative to Excel's. The kernel
            // feeds the values along the TOP ROW into what we call
            // column_input_cell (OOXML r1) -- a mapping validated against
            // Excel on hundreds of cells -- and Excel calls the cell that the
            // top row feeds its ROW input. So r1 is Excel's RowInput and r2 is
            // its ColumnInput, not the other way round.
            var r1 = ResolveCell(workbook, columnInput);
            var r2 = ResolveCell(workbook, rowInput);

            if (isTwoDimensional)
            {
                if (r1 == null || r2 == null)
                {
                    throw new InvalidOperationException(
                        "a two-variable table needs both input cells recorded");
                }

                full.Table(r1, r2);
                return;
            }

            if (r1 == null)
            {
                throw new InvalidOperationException("no input cell recorded for this table");
            }

            // One-variable: dtr says whether the single axis runs along a row
            // (feeding Excel's RowInput) or down a column (feeding ColumnInput).
            if (dtr)
            {
                full.Table(RowInput: r1);
            }
            else
            {
                full.Table(ColumnInput: r1);
            }
        }

        private static void RecordRestoreInfo(Excel.Worksheet metadata, DataTableRegionInfo region)
        {
            // Overwrite this table's live row if it has one, rather than
            // stacking another beside it.
            var existing = FindActiveMetadataRow(metadata, region.TableId);
            var row = existing > 0 ? existing : NextMetadataRow(metadata);
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
            ((Excel.Range)metadata.Cells[row, 9]).Value2 = region.Dtr ? 1d : 0d;
            ((Excel.Range)metadata.Cells[row, 10]).Value2 = "";
        }

        /// <summary>
        /// The row already describing this table, or 0. A table converted,
        /// restored and converted again used to append a row each time, and
        /// ReadOverrides handed every one of them to the engine -- so it
        /// received several generations of the same table at once, the older
        /// ones pointing at ranges that had since moved.
        /// </summary>
        private static int FindActiveMetadataRow(Excel.Worksheet metadata, string tableId)
        {
            var row = 2;
            while (true)
            {
                var candidate = CellText(metadata, row, 1);
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    return 0;
                }

                if (string.Equals(candidate, tableId, StringComparison.OrdinalIgnoreCase)
                    && string.IsNullOrWhiteSpace(CellText(metadata, row, 10)))
                {
                    return row;
                }

                row++;
            }
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
            ((Excel.Range)sheet.Cells[1, 9]).Value2 = "dtr";
            ((Excel.Range)sheet.Cells[1, 10]).Value2 = "restored_at";
            ((Excel.Range)sheet.Cells[1, 11]).Value2 = "created_live";
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

    internal sealed class CreateResult
    {
        public bool Created { get; set; }
        public int Rows { get; set; }
        public int Columns { get; set; }
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
