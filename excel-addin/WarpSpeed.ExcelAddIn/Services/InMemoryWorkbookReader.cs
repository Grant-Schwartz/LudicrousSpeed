using System;
using System.Collections.Generic;
using System.Globalization;
using WarpSpeed.ExcelAddIn.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace WarpSpeed.ExcelAddIn.Services
{
    /// <summary>
    /// Builds an <see cref="InlineWorkbook"/> directly from the live workbook
    /// over COM, as a faster alternative to <c>WorkbookSnapshotService</c>'s
    /// SaveCopyAs-to-disk-and-reimport path. Every worksheet is read with one
    /// bulk <c>Range.Formula</c> array read (not cell-by-cell, which is the
    /// well-known slow way to touch COM ranges), so cost scales with sheet
    /// size the same way native Excel recalculation does, rather than with
    /// per-cell COM marshalling overhead plus a full xlsx re-import.
    ///
    /// NOT YET VERIFIED AGAINST LIVE EXCEL. This was written and reasoned
    /// through carefully but could not be built or run in the environment
    /// that authored it (no Windows/Excel available there). Before relying
    /// on it: build on Windows, and run it against a real large workbook
    /// with dates, booleans, errors, merged cells, and both workbook- and
    /// sheet-scoped defined names, comparing its output against the existing
    /// SaveCopyAs path's result for the same workbook. See
    /// docs/windows-testing.md for the acceptance checklist.
    ///
    /// Two known gaps versus the file-based path, both by design for this
    /// first cut rather than oversights:
    ///  - Native Excel data table array formulas (the `{=TABLE(...)}` marker)
    ///    are not specially detected here the way xlsx_sanitize.rs detects
    ///    them from the OOXML. They get sent as plain formula text, which
    ///    IronCalc's parser currently accepts without erroring but almost
    ///    certainly won't evaluate usefully. Until this reader gains its own
    ///    detection (checking Range.HasArray / the `{=TABLE(` text shape),
    ///    workbooks with native data tables should keep using the file-based
    ///    path, which already handles them correctly.
    ///  - Only UsedRange per sheet is read, matching Excel's own notion of
    ///    "the populated area"; a sheet whose UsedRange was inflated by
    ///    formatting-only edits (common after cell deletes) will cost more
    ///    than strictly necessary but will not be incorrect.
    /// </summary>
    internal static class InMemoryWorkbookReader
    {
        public static InlineWorkbook Read(Excel.Workbook workbook)
        {
            var inline = new InlineWorkbook();

            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                inline.Sheets.Add(ReadSheet(worksheet));
            }

            foreach (Excel.Name name in workbook.Names)
            {
                var defined = TryReadDefinedName(name);
                if (defined != null)
                {
                    inline.DefinedNames.Add(defined);
                }
            }

            return inline;
        }

        private static InlineSheet ReadSheet(Excel.Worksheet worksheet)
        {
            var sheet = new InlineSheet { Name = worksheet.Name };

            Excel.Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int columnCount = usedRange.Columns.Count;

            // A truly empty sheet still reports a 1x1 UsedRange; skip it
            // rather than emitting one spurious empty cell.
            if (rowCount == 1 && columnCount == 1)
            {
                dynamic onlyCell = usedRange.Cells[1, 1];
                bool hasFormula = Convert.ToBoolean(onlyCell.HasFormula, CultureInfo.InvariantCulture);
                if (!hasFormula && onlyCell.Value2 == null)
                {
                    return sheet;
                }
            }

            int firstRow = usedRange.Row;
            int firstColumn = usedRange.Column;

            // dynamic (not a typed Excel.Range member access) to match the
            // rest of this codebase's COM usage, and because Formula (not
            // Formula2, which the pinned interop package version may not
            // even expose) is what every other file here already reads.
            // One bulk COM call for the whole used range. For a single cell,
            // Excel returns the scalar directly rather than a 2D array, so
            // that shape is normalized below alongside the array shape.
            dynamic dynamicRange = usedRange;
            object formulaBlock = dynamicRange.Formula;

            if (formulaBlock is object[,] formulas)
            {
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= columnCount; j++)
                    {
                        AddCellIfNotBlank(
                            sheet,
                            firstRow + i - 1,
                            firstColumn + j - 1,
                            formulas[i, j]);
                    }
                }
            }
            else
            {
                AddCellIfNotBlank(sheet, firstRow, firstColumn, formulaBlock);
            }

            return sheet;
        }

        private static void AddCellIfNotBlank(InlineSheet sheet, int row, int column, object rawValue)
        {
            var input = ConvertFormulaCellValue(rawValue);
            if (string.IsNullOrEmpty(input))
            {
                return;
            }

            sheet.Cells.Add(new InlineCell { Row = row, Column = column, Input = input });
        }

        /// <summary>
        /// Range.Formula returns formula text (already including the
        /// leading '=') for formula cells, and the literal typed value
        /// (double/string/bool/null) for plain cells -- mirroring
        /// WorkbookSnapshotService's ChangedCell.Input convention either way.
        /// </summary>
        private static string ConvertFormulaCellValue(object rawValue)
        {
            switch (rawValue)
            {
                case null:
                    return "";
                case string text:
                    // Formula text already starts with '='; a plain string
                    // value does not and is used as-is (IronCalc treats an
                    // unprefixed string input as a literal string value).
                    return text;
                case bool boolean:
                    // Excel/IronCalc boolean literals are unprefixed
                    // TRUE/FALSE, not .NET's "True"/"False".
                    return boolean ? "TRUE" : "FALSE";
                case double number:
                    return number.ToString("G17", CultureInfo.InvariantCulture);
                default:
                    return Convert.ToString(rawValue, CultureInfo.InvariantCulture) ?? "";
            }
        }

        private static InlineDefinedName? TryReadDefinedName(Excel.Name name)
        {
            try
            {
                var fullName = Convert.ToString(name.Name, CultureInfo.InvariantCulture) ?? "";
                var refersTo = Convert.ToString(name.RefersTo, CultureInfo.InvariantCulture) ?? "";
                if (string.IsNullOrEmpty(fullName) || string.IsNullOrEmpty(refersTo))
                {
                    return null;
                }

                // Sheet-scoped names come back from Workbook.Names as
                // "'Sheet Name'!ShortName" or "SheetName!ShortName"; parse
                // the qualifier out of the string itself rather than relying
                // on Name.Parent (not confident that reliably distinguishes
                // sheet- from workbook-scope across Excel versions -- string
                // parsing is the documented, version-stable way names round
                // trip through this property).
                string shortName = fullName;
                string? scopeSheetName = null;
                var bangIndex = fullName.IndexOf('!');
                if (bangIndex > 0)
                {
                    var qualifier = fullName.Substring(0, bangIndex);
                    if (qualifier.Length >= 2 && qualifier[0] == '\'' && qualifier[qualifier.Length - 1] == '\'')
                    {
                        qualifier = qualifier.Substring(1, qualifier.Length - 2).Replace("''", "'");
                    }
                    scopeSheetName = qualifier;
                    shortName = fullName.Substring(bangIndex + 1);
                }

                return new InlineDefinedName
                {
                    Name = shortName,
                    ScopeSheetName = scopeSheetName,
                    Formula = refersTo,
                };
            }
            catch (Exception)
            {
                // Some built-in/hidden names (print areas, filter databases,
                // etc.) throw on RefersTo or have shapes that don't fit this
                // model; skip rather than aborting the whole read.
                return null;
            }
        }
    }
}
