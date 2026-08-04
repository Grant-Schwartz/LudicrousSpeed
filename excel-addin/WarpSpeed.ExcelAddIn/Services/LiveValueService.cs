using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;

namespace WarpSpeed.ExcelAddIn.Services
{
    /// <summary>
    /// Backs the WS.LIVE worksheet function with values computed by the Rust
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

        public static CellObservable GetOrCreate(string normalizedAddress)
        {
            return Cells.GetOrAdd(normalizedAddress, _ => new CellObservable());
        }

        /// <summary>
        /// Pushes a batch of engine results to any WS.LIVE cells watching
        /// them. Addresses that nothing is watching are still retained, so a
        /// WS.LIVE formula added later picks up the last known value on its
        /// first Subscribe rather than sitting at #N/A until the next run.
        /// </summary>
        public static void PublishAll(IEnumerable<KeyValuePair<string, object>> values)
        {
            foreach (var entry in values)
            {
                GetOrCreate(NormalizeAddress(entry.Key)).Publish(entry.Value);
            }
        }

        public static int WatchedCellCount => Cells.Count;
    }

    /// <summary>
    /// One live cell. Excel-DNA subscribes when a WS.LIVE formula is first
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
                // Replay the last known value so a newly-added WS.LIVE cell
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
    /// <c>=WS.LIVE("LBO (Share Price)!K232")</c> displays whatever the engine
    /// last published for that address, and updates in place when a new value
    /// arrives -- without Excel evaluating anything for it.
    /// </summary>
    public static class LiveValueFunctions
    {
        [ExcelFunction(
            Name = "WS.LIVE",
            Description = "Displays the value WarpSpeed computed for a model cell, updated live.")]
        public static object WsLive(
            [ExcelArgument(Name = "address", Description = "Cell to mirror, e.g. \"Sheet1!A1\"")]
            string address)
        {
            var key = LiveValueService.NormalizeAddress(address);
            if (key.Length == 0)
            {
                return ExcelError.ExcelErrorValue;
            }

            return ExcelAsyncUtil.Observe(
                "WS.LIVE",
                key,
                () => LiveValueService.GetOrCreate(key));
        }

        [ExcelFunction(
            Name = "WS.LIVECOUNT",
            Description = "Number of cell addresses WarpSpeed is currently tracking for live values.")]
        public static object WsLiveCount()
        {
            return (double)LiveValueService.WatchedCellCount;
        }
    }
}
