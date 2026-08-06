using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using LudicrousSpeed.ExcelAddIn.Models;

namespace LudicrousSpeed.ExcelAddIn.Services
{
    /// <summary>
    /// Delivers engine results to the sheet by publishing them to
    /// <see cref="LiveValueService"/>, which pushes them into WS.LIVE cells
    /// over RTD.
    ///
    /// This replaced a much larger writer that tried to set a formula cell's
    /// cached value in place. That is not possible in Excel: Range.Value2
    /// replaces the formula, the XLL C API's xlSet is macro-sheet-only and
    /// silently no-ops on a worksheet cell, and setting Value2 then restoring
    /// the formula re-evaluates it. All three were tested against live Excel
    /// and failed, each for a different reason, so the probe and the
    /// apply/verify/restore machinery that depended on them are gone.
    /// </summary>
    internal sealed class LiveValuePublisher
    {
        /// <summary>
        /// Pushes every value the engine produced -- ordinary formula cells
        /// and data table outputs alike -- to any live cells watching those
        /// addresses. Returns how many were published.
        /// </summary>
        public int Publish(EngineResponse response)
        {
            if (!response.Ok || response.Result == null)
            {
                return 0;
            }

            var plan = response.Result.Writeback;
            var published = new List<KeyValuePair<string, object>>(
                plan.Cells.Count + plan.DataTableCells.Count);

            foreach (var candidate in plan.Cells)
            {
                if (TryConvertValue(candidate.ValueKind, candidate.Value, out var value))
                {
                    published.Add(new KeyValuePair<string, object>(
                        candidate.SheetName + "!" + candidate.Address, value));
                }
            }

            // Data table outputs are the highest-value cells to drive: each
            // one otherwise costs Excel a full re-evaluation of the table's
            // source formula cone.
            foreach (var candidate in plan.DataTableCells)
            {
                if (TryConvertValue(candidate.ValueKind, candidate.Value, out var value))
                {
                    published.Add(new KeyValuePair<string, object>(
                        candidate.SheetName + "!" + candidate.Address, value));
                }
            }

            LiveValueService.PublishAll(published);
            return published.Count;
        }

        private static bool TryConvertValue(string valueKind, JToken? token, out object value)
        {
            value = "";
            try
            {
                switch (valueKind.Trim().ToLowerInvariant())
                {
                    case "number":
                        value = token?.ToObject<double>() ?? 0.0;
                        return true;
                    case "string":
                        value = token?.ToObject<string>() ?? "";
                        return true;
                    case "boolean":
                        value = token?.ToObject<bool>() ?? false;
                        return true;
                    default:
                        // Blanks and error values are deliberately not
                        // published: there is nothing useful to display, and
                        // an error should never overwrite a number someone
                        // might act on.
                        return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
