using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using LudicrousSpeed.ExcelAddIn.Models;

namespace LudicrousSpeed.ExcelAddIn.Services
{
    internal sealed class WorkbookChangeTracker
    {
        public const int DirtyCellLimit = 10000;
        private const string ReportSheetName = "_LudicrousSpeed_Report";
        private const string FallbackSheetName = "_LudicrousSpeed_Fallbacks";
        internal const string DataTableSheetName = "_LudicrousSpeed_DataTables";

        private readonly object gate = new object();
        private readonly Dictionary<string, WorkbookDirtyState> states =
            new Dictionary<string, WorkbookDirtyState>(StringComparer.OrdinalIgnoreCase);

        private Excel.Application? application;
        private bool subscribed;
        private int suspendCount;

        public void Start()
        {
            if (subscribed)
            {
                return;
            }

            application = (Excel.Application)ExcelDnaUtil.Application;
            application.SheetChange += OnSheetChange;
            subscribed = true;
        }

        public void Stop()
        {
            if (!subscribed || application == null)
            {
                return;
            }

            application.SheetChange -= OnSheetChange;
            subscribed = false;
            application = null;
        }

        public IDisposable SuspendTracking()
        {
            lock (gate)
            {
                suspendCount++;
            }

            return new TrackingSuspension(this);
        }

        private void ResumeTracking()
        {
            lock (gate)
            {
                if (suspendCount > 0)
                {
                    suspendCount--;
                }
            }
        }

        public WorkbookChangeSet CaptureForSnapshot(object workbookObject)
        {
            var workbookId = GetWorkbookId(workbookObject);
            var sheetSignature = GetSheetSignature(workbookObject);

            lock (gate)
            {
                var state = GetOrCreateState(workbookId);
                var topologyChanged = state.IsWarm
                    && !string.Equals(state.LastSheetSignature, sheetSignature, StringComparison.Ordinal);

                return new WorkbookChangeSet
                {
                    WorkbookId = workbookId,
                    SheetSignature = sheetSignature,
                    IsWarm = state.IsWarm,
                    ForceReload = state.ForceReload || topologyChanged,
                    DirtyCells = new List<DirtyCellAddress>(state.DirtyCells.Values),
                };
            }
        }

        public void MarkRunSucceeded(WorkbookSnapshot snapshot)
        {
            if (string.IsNullOrWhiteSpace(snapshot.WorkbookId))
            {
                return;
            }

            lock (gate)
            {
                var state = GetOrCreateState(snapshot.WorkbookId);
                state.IsWarm = true;
                state.ForceReload = false;
                state.LastSheetSignature = snapshot.SheetSignature;
                state.DirtyCells.Clear();
            }
        }

        public static string GetWorkbookId(object workbookObject)
        {
            var workbook = (Excel.Workbook)workbookObject;
            var name = Convert.ToString(workbook.Name, CultureInfo.InvariantCulture) ?? "Workbook";
            var path = Convert.ToString(workbook.Path, CultureInfo.InvariantCulture) ?? "";
            if (!string.IsNullOrWhiteSpace(path))
            {
                var fullName = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(fullName))
                {
                    return fullName;
                }
            }

            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}:{1}",
                name,
                RuntimeHelpers.GetHashCode(workbook));
        }

        private static string GetSheetSignature(object workbookObject)
        {
            var workbook = (Excel.Workbook)workbookObject;
            var names = new List<string>();
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                var name = Convert.ToString(worksheet.Name, CultureInfo.InvariantCulture) ?? "";
                if (IsLudicrousSpeedReportSheet(name))
                {
                    continue;
                }

                names.Add(name);
            }

            return string.Join("\u001f", names);
        }

        private void OnSheetChange(object sheetObject, Excel.Range target)
        {
            try
            {
                lock (gate)
                {
                    if (suspendCount > 0)
                    {
                        return;
                    }
                }

                var worksheet = sheetObject as Excel.Worksheet;
                if (worksheet == null || target == null)
                {
                    return;
                }

                var sheetName = Convert.ToString(worksheet.Name, CultureInfo.InvariantCulture) ?? "";
                if (IsLudicrousSpeedReportSheet(sheetName))
                {
                    return;
                }

                var workbook = (Excel.Workbook)worksheet.Parent;
                var workbookId = GetWorkbookId(workbook);
                var count = GetRangeCellCount(target);

                lock (gate)
                {
                    var state = GetOrCreateState(workbookId);
                    if (state.ForceReload || count > DirtyCellLimit || state.DirtyCells.Count + count > DirtyCellLimit)
                    {
                        state.ForceReload = true;
                        state.DirtyCells.Clear();
                        return;
                    }

                    foreach (Excel.Range cell in target.Cells)
                    {
                        var row = Convert.ToInt32(cell.Row, CultureInfo.InvariantCulture);
                        var column = Convert.ToInt32(cell.Column, CultureInfo.InvariantCulture);
                        var address = Convert.ToString(
                            cell.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing),
                            CultureInfo.InvariantCulture) ?? "";
                        var key = string.Format(CultureInfo.InvariantCulture, "{0}!{1}:{2}", sheetName, row, column);
                        state.DirtyCells[key] = new DirtyCellAddress
                        {
                            SheetName = sheetName,
                            Row = row,
                            Column = column,
                            Address = address,
                        };
                    }
                }
            }
            catch
            {
                // If event tracking becomes uncertain, the next snapshot will reload from the workbook file.
                try
                {
                    var worksheet = sheetObject as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        return;
                    }

                    var workbookId = GetWorkbookId((Excel.Workbook)worksheet.Parent);
                    lock (gate)
                    {
                        var state = GetOrCreateState(workbookId);
                        state.ForceReload = true;
                        state.DirtyCells.Clear();
                    }
                }
                catch
                {
                    // Nothing else to do without a reliable workbook identity.
                }
            }
        }

        private sealed class TrackingSuspension : IDisposable
        {
            private readonly WorkbookChangeTracker owner;
            private bool disposed;

            public TrackingSuspension(WorkbookChangeTracker owner)
            {
                this.owner = owner;
            }

            public void Dispose()
            {
                if (disposed)
                {
                    return;
                }

                disposed = true;
                owner.ResumeTracking();
            }
        }

        private static long GetRangeCellCount(Excel.Range range)
        {
            try
            {
                return Convert.ToInt64(range.CountLarge, CultureInfo.InvariantCulture);
            }
            catch
            {
                return Convert.ToInt64(range.Count, CultureInfo.InvariantCulture);
            }
        }

        private WorkbookDirtyState GetOrCreateState(string workbookId)
        {
            if (!states.TryGetValue(workbookId, out var state))
            {
                state = new WorkbookDirtyState();
                states[workbookId] = state;
            }

            return state;
        }

        /// <summary>
        /// Sheets LudicrousSpeed writes to itself. Edits here must never be
        /// reported as workbook changes: they aren't model edits, and a cached
        /// engine model won't contain a sheet LudicrousSpeed created after the last
        /// snapshot -- which previously failed the whole run with "sheet not
        /// found for changed cell".
        /// </summary>
        private static bool IsLudicrousSpeedReportSheet(string sheetName)
        {
            return string.Equals(sheetName, ReportSheetName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(sheetName, FallbackSheetName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(sheetName, DataTableSheetName, StringComparison.OrdinalIgnoreCase);
        }
    }

    internal sealed class WorkbookChangeSet
    {
        public string WorkbookId { get; set; } = "";

        public string SheetSignature { get; set; } = "";

        public bool IsWarm { get; set; }

        public bool ForceReload { get; set; }

        public List<DirtyCellAddress> DirtyCells { get; set; } = new List<DirtyCellAddress>();
    }

    internal sealed class DirtyCellAddress
    {
        public string SheetName { get; set; } = "";

        public int Row { get; set; }

        public int Column { get; set; }

        public string Address { get; set; } = "";
    }

    internal sealed class WorkbookDirtyState
    {
        public bool IsWarm { get; set; }

        public bool ForceReload { get; set; }

        public string? LastSheetSignature { get; set; }

        public Dictionary<string, DirtyCellAddress> DirtyCells { get; } =
            new Dictionary<string, DirtyCellAddress>(StringComparer.OrdinalIgnoreCase);
    }
}
