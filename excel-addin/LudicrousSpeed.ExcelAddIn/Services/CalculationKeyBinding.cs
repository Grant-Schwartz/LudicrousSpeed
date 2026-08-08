using System;
using ExcelDna.Integration;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace LudicrousSpeed.ExcelAddIn.Services
{
    /// <summary>
    /// Routes F9 to the LudicrousSpeed engine instead of Excel's calculation.
    ///
    /// NOT YET VERIFIED AGAINST LIVE EXCEL, and off by default, because of one
    /// specific risk: F9 has two unrelated jobs in Excel. Outside edit mode it
    /// recalculates, which is the one we want. Inside the formula bar, with a
    /// sub-expression selected, it evaluates just that fragment in place -- a
    /// keystroke M&amp;A analysts use constantly to read a term out of a long
    /// formula. Application.OnKey is documented to assign a procedure to a
    /// keystroke, and macros cannot run while a cell is being edited, so the
    /// in-formula behavior is expected to survive untouched. "Expected" is not
    /// "observed": this was written without Windows or Excel to test against.
    /// Confirm the formula-bar case by hand before enabling this for anyone
    /// who edits formulas for a living -- see docs/windows-testing.md.
    ///
    /// Only plain F9 is taken. Shift+F9, Ctrl+Alt+F9 and Ctrl+Alt+Shift+F9 stay
    /// native deliberately, so Excel's own answer is always one keystroke away
    /// for comparison, and so no workbook can become uncheckable.
    ///
    /// The binding is per-Excel-session and per-application, not per-workbook.
    /// Excel applies OnKey to every open workbook, including ones the engine
    /// has never seen, so the handler falls through to a native calculation
    /// whenever it cannot run the engine.
    /// </summary>
    public static class CalculationKeyBinding
    {
        /// <summary>
        /// The name Excel-DNA registers <see cref="LudicrousSpeedCalculateNow"/>
        /// under, and the name OnKey resolves to find it.
        /// </summary>
        private const string MacroName = "LudicrousSpeedCalculateNow";

        private const string KeySequence = "{F9}";

        private const string SettingsKey = @"Software\LudicrousSpeed";
        private const string SettingsValue = "InterceptF9";

        private static Action? recalculate;
        private static bool bound;
        private static bool running;

        public static bool Enabled { get; private set; }

        /// <summary>
        /// Called once at add-in load. Restores the persisted preference and
        /// binds the key if it was left on.
        /// </summary>
        public static void Initialize(Action recalculateHandler)
        {
            recalculate = recalculateHandler;
            Enabled = ReadPersistedSetting();
            if (Enabled)
            {
                Bind();
            }
        }

        public static void SetEnabled(bool enabled)
        {
            Enabled = enabled;
            WritePersistedSetting(enabled);

            if (enabled)
            {
                Bind();
            }
            else
            {
                Unbind();
            }
        }

        /// <summary>
        /// Called at add-in unload. OnKey assignments do not survive an Excel
        /// session, so a crash cannot leave F9 permanently bound to a macro
        /// that no longer exists -- but an orderly unload should still hand the
        /// key back rather than rely on that.
        /// </summary>
        public static void Shutdown()
        {
            Unbind();
        }

        private static void Bind()
        {
            if (bound)
            {
                return;
            }

            try
            {
                var excel = (Excel.Application)ExcelDnaUtil.Application;
                excel.OnKey(KeySequence, MacroName);
                bound = true;
            }
            catch
            {
                // Leave F9 native rather than half-bound.
                bound = false;
            }
        }

        private static void Unbind()
        {
            if (!bound)
            {
                return;
            }

            try
            {
                var excel = (Excel.Application)ExcelDnaUtil.Application;
                // Omitting the procedure argument restores Excel's default
                // handling. Passing "" would instead make F9 do nothing at
                // all, which is the one outcome worse than either behavior.
                excel.OnKey(KeySequence);
            }
            catch
            {
                // Nothing useful to do; the binding dies with the session.
            }
            finally
            {
                bound = false;
            }
        }

        /// <summary>
        /// The macro F9 is bound to. Excel-DNA registers this as an XLL
        /// command because it is public, static, and returns void.
        ///
        /// Every failure path here ends in a native Excel calculation. F9 must
        /// never become a dead key: a user who presses it and sees nothing
        /// happen has no way to tell a broken add-in from an up-to-date model.
        /// </summary>
        [ExcelCommand(Name = MacroName)]
        public static void LudicrousSpeedCalculateNow()
        {
            // F9 is held down, or the engine run triggered a calculation that
            // somehow re-entered. Swallow rather than stack up engine runs.
            if (running)
            {
                return;
            }

            running = true;
            try
            {
                if (recalculate == null || !HasActiveWorkbook())
                {
                    CalculateNatively();
                    return;
                }

                recalculate();
            }
            catch
            {
                CalculateNatively();
            }
            finally
            {
                running = false;
            }
        }

        private static bool HasActiveWorkbook()
        {
            try
            {
                var excel = (Excel.Application)ExcelDnaUtil.Application;
                return excel.ActiveWorkbook != null;
            }
            catch
            {
                return false;
            }
        }

        private static void CalculateNatively()
        {
            try
            {
                dynamic excel = ExcelDnaUtil.Application;
                excel.Calculate();
            }
            catch
            {
                // If even this fails, Excel is in a state the add-in cannot
                // improve on.
            }
        }

        private static bool ReadPersistedSetting()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(SettingsKey);
                return key?.GetValue(SettingsValue) as string == "1";
            }
            catch
            {
                return false;
            }
        }

        private static void WritePersistedSetting(bool enabled)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(SettingsKey);
                key?.SetValue(SettingsValue, enabled ? "1" : "0", RegistryValueKind.String);
            }
            catch
            {
                // The preference degrades to session-only. Not worth a dialog.
            }
        }
    }
}
