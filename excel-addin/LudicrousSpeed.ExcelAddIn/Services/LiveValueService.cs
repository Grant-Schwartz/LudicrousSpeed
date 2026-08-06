using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;

namespace LudicrousSpeed.ExcelAddIn.Services
{
    /// <summary>
    /// Backs the LS.LIVE worksheet function with values computed by the Rust
    /// engine, pushed into Excel over RTD.
    ///
    /// WHY RTD: writing a computed value into a cell that already holds a
    /// formula is not possible through Excel's object model -- Range.Value2
    /// replaces the formula, the XLL C API's xlSet is macro-sheet-only and
    /// silently no-ops on worksheet cells, and setting Value2 then restoring
    /// the formula re-triggers evaluation of that cell. All three were tested
    /// against live Excel and failed. RTD is the one mechanism Excel
    /// sanctions for "an external process supplies this cell's value", which
    /// is why every market-data vendor uses it.
    ///
    /// WHY IT MAKES THE WHOLE WORKBOOK FAST, not just these cells: an RTD
    /// value landing marks its dependents dirty, so Excel recalculates
    /// downstream normally. The cells worth wiring up are therefore not the
    /// outputs a user reads, but the places Excel does *repeated* work --
    /// data table scenario cells (Excel re-evaluates the source formula's
    /// cone once per scenario) and circular-component break points (Excel
    /// iterates the region to convergence). Injecting there removes the
    /// multiplicative cost and leaves Excel a single linear pass. Injecting
    /// into an ordinary formula in a linear chain yields nothing, because
    /// Excel still has to evaluate that cell's precedents regardless.
    ///
    /// This uses Excel-DNA's built-in observer RTD server via
    /// ExcelAsyncUtil.Observe rather than a hand-written IRtdServer, so
    /// there is no COM registration or ProgId of our own to get wrong.
    ///
    /// NOT YET VERIFIED AGAINST LIVE EXCEL.
    /// </summary>
    internal static class LiveValueService
    {
        private static readonly ConcurrentDictionary<string, CellObservable> Cells =
            new ConcurrentDictionary<string, CellObservable>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Canonical key for a cell, so the UDF's argument and the engine's
        /// published address agree regardless of spacing/casing. Sheet names
        /// may contain '!' only when quoted, so split on the last one.
        /// </summary>
        public static string NormalizeAddress(string address)
        {
            var trimmed = (address ?? "").Trim();
            var bang = trimmed.LastIndexOf('!');
            if (bang <= 0)
            {
                return trimmed.ToUpperInvariant();
            }

            var sheet = trimmed.Substring(0, bang).Trim().Trim('\'').Replace("''", "'");
            var cell = trimmed.Substring(bang + 1).Trim().Replace("$", "");
            return (sheet + "!" + cell).ToUpperInvariant();
        }

        /// <summary>
        /// Builds a registry key from a worksheet reference, so
        /// <c>=LS.LIVE(E2)</c> works the way an Excel user expects rather than
        /// requiring the sheet-qualified string form. Excel hands the UDF an
        /// ExcelReference (zero-indexed) only when the argument is declared
        /// AllowReference; xlSheetNm resolves it to "[Book.xlsx]Sheet", whose
        /// workbook prefix is dropped since the registry is keyed by sheet.
        /// </summary>
        public static string AddressFromReference(ExcelReference reference)
        {
            try
            {
                var sheetName =
                    Convert.ToString(XlCall.Excel(XlCall.xlSheetNm, reference), CultureInfo.InvariantCulture)
                    ?? "";
                var bracket = sheetName.LastIndexOf(']');
                if (bracket >= 0)
                {
                    sheetName = sheetName.Substring(bracket + 1);
                }

                if (sheetName.Length == 0)
                {
                    return "";
                }

                var column = ColumnLetters(reference.ColumnFirst + 1);
                var row = reference.RowFirst + 1;
                return NormalizeAddress(sheetName + "!" + column + row.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception)
            {
                return "";
            }
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

        public static CellObservable GetOrCreate(string normalizedAddress)
        {
            return Cells.GetOrAdd(normalizedAddress, _ => new CellObservable());
        }

        /// <summary>
        /// Pushes a batch of engine results to any LS.LIVE cells watching
        /// them. Addresses that nothing is watching are still retained, so a
        /// LS.LIVE formula added later picks up the last known value on its
        /// first Subscribe rather than sitting at #N/A until the next run.
        /// </summary>
        public static void PublishAll(IEnumerable<KeyValuePair<string, object>> values)
        {
            var published = 0;
            foreach (var entry in values)
            {
                GetOrCreate(NormalizeAddress(entry.Key)).Publish(entry.Value);
                published++;
            }

            if (published == 0)
            {
                return;
            }

            // Not every cell is publishable: cells that evaluate to an error
            // (e.g. functions IronCalc doesn't implement, like XIRR), cells
            // inside or downstream of a fallback region, and data table
            // outputs are all deliberately excluded from the writeback-safe
            // set. Without this, a LS.LIVE formula pointed at one of those
            // sits at #N/A forever, indistinguishable from "still waiting for
            // the first run" -- which is a genuinely confusing failure.
            // Anything still valueless after a publish cycle gets told so.
            // Only ever touches cells that have never received a value, so a
            // real value from an earlier run is never clobbered.
            foreach (var cell in Cells.Values)
            {
                cell.PublishIfNeverSet(NotPublishedMessage);
            }
        }

        internal const string NotPublishedMessage =
            "#not-published - LudicrousSpeed has no trusted value for this cell "
            + "(it evaluates to an error, or sits in/downstream of a fallback region)";

        public static int WatchedCellCount => Cells.Count;

        /// <summary>
        /// Diagnostic: describes the registry state for one address, so a
        /// #N/A can be traced to the actual broken link -- nothing published
        /// at all, published under a different key, or published but never
        /// delivered to a subscriber.
        /// </summary>
        public static string Describe(string address)
        {
            var key = NormalizeAddress(address);
            if (!Cells.TryGetValue(key, out var cell))
            {
                return $"key=\"{key}\" NOT in registry (total tracked={Cells.Count})";
            }

            return $"key=\"{key}\" {cell.Describe()} (total tracked={Cells.Count})";
        }

        /// <summary>
        /// Diagnostic: a sample of keys actually present, to eyeball against
        /// what a LS.LIVE formula is asking for.
        /// </summary>
        public static string SampleKeys(int count)
        {
            var keys = new List<string>();
            foreach (var key in Cells.Keys)
            {
                keys.Add(key);
                if (keys.Count >= count)
                {
                    break;
                }
            }

            return keys.Count == 0 ? "<registry empty>" : string.Join(" | ", keys);
        }
    }

    /// <summary>
    /// One live cell. Excel-DNA subscribes when a LS.LIVE formula is first
    /// evaluated and unsubscribes when the last such formula goes away.
    /// Only IExcelObservable is implemented here -- IExcelObserver is
    /// supplied by Excel-DNA and we only ever call OnNext on it.
    /// </summary>
    internal sealed class CellObservable : IExcelObservable
    {
        private readonly object gate = new object();
        private readonly List<IExcelObserver> observers = new List<IExcelObserver>();
        private object? latest;
        private bool hasValue;

        public IDisposable Subscribe(IExcelObserver observer)
        {
            lock (gate)
            {
                observers.Add(observer);
                // Replay the last known value so a newly-added LS.LIVE cell
                // shows a number immediately instead of #N/A until the next
                // engine run.
                if (hasValue)
                {
                    observer.OnNext(latest);
                }
            }

            return new Unsubscriber(this, observer);
        }

        public void Publish(object value)
        {
            List<IExcelObserver> snapshot;
            lock (gate)
            {
                latest = value;
                hasValue = true;
                snapshot = new List<IExcelObserver>(observers);
            }

            // Outside the lock: OnNext marshals into Excel and shouldn't be
            // held up by, or deadlock against, another thread publishing.
            foreach (var observer in snapshot)
            {
                observer.OnNext(value);
            }
        }

        /// <summary>
        /// Publishes <paramref name="value"/> only if this cell has never
        /// received one, so an explanatory marker can be delivered to
        /// never-published addresses without overwriting real results.
        /// </summary>
        public void PublishIfNeverSet(object value)
        {
            lock (gate)
            {
                if (hasValue)
                {
                    return;
                }
            }

            Publish(value);
        }

        internal string Describe()
        {
            lock (gate)
            {
                var value = hasValue
                    ? Convert.ToString(latest, CultureInfo.InvariantCulture) ?? "<null>"
                    : "<none published>";
                return $"observers={observers.Count} hasValue={hasValue} value={value}";
            }
        }

        private void Remove(IExcelObserver observer)
        {
            lock (gate)
            {
                observers.Remove(observer);
            }
        }

        private sealed class Unsubscriber : IDisposable
        {
            private readonly CellObservable owner;
            private readonly IExcelObserver observer;
            private bool disposed;

            public Unsubscriber(CellObservable owner, IExcelObserver observer)
            {
                this.owner = owner;
                this.observer = observer;
            }

            public void Dispose()
            {
                if (disposed)
                {
                    return;
                }

                disposed = true;
                owner.Remove(observer);
            }
        }
    }

    /// <summary>
    /// The worksheet-facing surface. A cell whose formula is
    /// <c>=LS.LIVE("LBO (Share Price)!K232")</c> displays whatever the engine
    /// last published for that address, and updates in place when a new value
    /// arrives -- without Excel evaluating anything for it.
    /// </summary>
    public static class LiveValueFunctions
    {
        [ExcelFunction(
            Name = "LS.LIVE",
            Description = "Displays the value LudicrousSpeed computed for a model cell, updated live.")]
        public static object LsLive(
            [ExcelArgument(
                Name = "cell",
                Description = "Cell to mirror: a reference like E2, or text like \"Sheet1!E2\"",
                AllowReference = true)]
            object cell)
        {
            // AllowReference means an unquoted argument arrives as an
            // ExcelReference rather than the referenced cell's *value*, which
            // is what =LS.LIVE(E2) needs -- without it Excel passes E2's
            // contents and the lookup is for an address named "45378".
            string key;
            if (cell is ExcelReference reference)
            {
                key = LiveValueService.AddressFromReference(reference);
                if (key.Length == 0)
                {
                    return ExcelError.ExcelErrorRef;
                }
            }
            else
            {
                key = LiveValueService.NormalizeAddress(
                    Convert.ToString(cell, CultureInfo.InvariantCulture) ?? "");
            }

            if (key.Length == 0)
            {
                return ExcelError.ExcelErrorValue;
            }

            return ExcelAsyncUtil.Observe(
                "LS.LIVE",
                key,
                () => LiveValueService.GetOrCreate(key));
        }


        [ExcelFunction(
            Name = "LS.LIVECOUNT",
            Description = "Number of cell addresses LudicrousSpeed is currently tracking for live values.",
            IsVolatile = true)]
        public static object LsLiveCount()
        {
            return (double)LiveValueService.WatchedCellCount;
        }

        [ExcelFunction(
            Name = "LS.LIVEDEBUG",
            Description = "Diagnostic: registry state for one address (is it present, does it have a value, is anything subscribed).",
            IsVolatile = true)]
        public static object LsLiveDebug(
            [ExcelArgument(Name = "address", Description = "Cell to inspect, e.g. \"Sheet1!A1\"")]
            string address)
        {
            return LiveValueService.Describe(address);
        }

        [ExcelFunction(
            Name = "LS.LIVEKEYS",
            Description = "Diagnostic: sample of addresses currently in the live-value registry.",
            IsVolatile = true)]
        public static object LsLiveKeys(
            [ExcelArgument(Name = "count", Description = "How many keys to show")] object count)
        {
            var n = 5;
            if (count is double d && d >= 1)
            {
                n = (int)d;
            }

            return LiveValueService.SampleKeys(Math.Min(n, 40));
        }
    }
}
