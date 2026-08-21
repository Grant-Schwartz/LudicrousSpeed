using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration;
using Newtonsoft.Json;
using LudicrousSpeed.ExcelAddIn.Models;

namespace LudicrousSpeed.ExcelAddIn.Interop
{
    internal sealed class NativeEngineClient
    {
        private const string WindowsDll = "ludicrous_engine.dll";

        private static bool nativeLibraryResolved;
        private static string nativeLibraryPath = "";

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr LoadLibraryW(string fileName);

        /// <summary>
        /// Loads the engine from the add-in's own folder, before anything
        /// P/Invokes into it.
        ///
        /// DllImport by bare name resolves through Windows' standard search
        /// order: Excel's own directory, the current working directory, then
        /// PATH. The folder the .xll was loaded from is not in that list. So
        /// the engine sitting directly beside the add-in was found only when
        /// Excel's working directory happened to be that folder -- true right
        /// after browsing to it in the add-ins dialog, false after a restart,
        /// which is why reinstalling appeared to fix it and never lasted.
        ///
        /// Loading it once by absolute path puts the module in the process. The
        /// later DllImport matches on base name against what is already loaded
        /// and never searches at all.
        /// </summary>
        public static void EnsureNativeLibraryLoaded()
        {
            if (nativeLibraryResolved)
            {
                return;
            }

            nativeLibraryResolved = true;
            try
            {
                var addInFolder = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
                if (string.IsNullOrEmpty(addInFolder))
                {
                    return;
                }

                var candidate = Path.Combine(addInFolder, WindowsDll);
                if (!File.Exists(candidate))
                {
                    return;
                }

                if (LoadLibraryW(candidate) != IntPtr.Zero)
                {
                    nativeLibraryPath = candidate;
                }
            }
            catch (Exception)
            {
                // Falls back to the normal search order, which is where this
                // started. Never worth failing add-in startup over.
            }
        }

        public EngineResponse Run(WorkbookSnapshot snapshot, out long nativeCallMs)
        {
            var payload = JsonConvert.SerializeObject(snapshot);
            var responsePtr = IntPtr.Zero;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                responsePtr = ludicrous_run_json(payload);
                stopwatch.Stop();
                nativeCallMs = stopwatch.ElapsedMilliseconds;
                if (responsePtr == IntPtr.Zero)
                {
                    return EngineResponse.Failed("Rust engine returned a null response.");
                }

                var json = PtrToUtf8String(responsePtr);
                return JsonConvert.DeserializeObject<EngineResponse>(json)
                    ?? EngineResponse.Failed("Rust engine returned an empty response.");
            }
            catch (DllNotFoundException)
            {
                stopwatch.Stop();
                nativeCallMs = stopwatch.ElapsedMilliseconds;
                var looked = string.IsNullOrEmpty(nativeLibraryPath)
                    ? Path.GetDirectoryName(ExcelDnaUtil.XllPath) ?? "the add-in folder"
                    : nativeLibraryPath;
                return EngineResponse.Failed(
                    $"Could not load {WindowsDll}. Expected it beside the add-in, at {looked}. "
                    + "Reinstalling with Install.cmd puts it there.");
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                nativeCallMs = stopwatch.ElapsedMilliseconds;
                return EngineResponse.Failed(ex.Message);
            }
            finally
            {
                if (responsePtr != IntPtr.Zero)
                {
                    ludicrous_free_string(responsePtr);
                }
            }
        }

        private static string PtrToUtf8String(IntPtr ptr)
        {
            var length = 0;
            while (Marshal.ReadByte(ptr, length) != 0)
            {
                length++;
            }

            var buffer = new byte[length];
            Marshal.Copy(ptr, buffer, 0, length);
            return Encoding.UTF8.GetString(buffer);
        }

        /// <summary>
        /// Reads how far the current run has got. Safe to call while
        /// <see cref="Run"/> is in flight on another thread -- that is the
        /// point of it, since Run blocks for the length of the calculation.
        ///
        /// Returns an idle snapshot rather than throwing if the native library
        /// is missing or too old to export the symbol. A progress indicator is
        /// a courtesy; failing to read one must never take down a run that is
        /// otherwise working.
        /// </summary>
        public EngineProgress ReadProgress()
        {
            try
            {
                ludicrous_progress(out var phase, out var done, out var total);
                return new EngineProgress((EnginePhase)phase, done, total);
            }
            catch (DllNotFoundException)
            {
                return default;
            }
            catch (EntryPointNotFoundException)
            {
                return default;
            }
        }

        [DllImport(WindowsDll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern IntPtr ludicrous_run_json(string requestJson);

        [DllImport(WindowsDll, CallingConvention = CallingConvention.Cdecl)]
        private static extern void ludicrous_free_string(IntPtr value);

        [DllImport(WindowsDll, CallingConvention = CallingConvention.Cdecl)]
        private static extern void ludicrous_progress(out uint phase, out ulong done, out ulong total);
    }

    /// <summary>Mirrors the PHASE_* constants in the engine's progress module.</summary>
    internal enum EnginePhase : uint
    {
        Idle = 0,
        Loading = 1,
        Analyzing = 2,
        Evaluating = 3,
        DataTables = 4,
    }

    internal readonly struct EngineProgress
    {
        public EngineProgress(EnginePhase phase, ulong done, ulong total)
        {
            Phase = phase;
            Done = done;
            Total = total;
        }

        public EnginePhase Phase { get; }

        public ulong Done { get; }

        /// <summary>Zero means the phase is indeterminate -- show a name, not a percentage.</summary>
        public ulong Total { get; }

        public bool IsRunning => Phase != EnginePhase.Idle;

        /// <summary>
        /// Status bar text for this snapshot, in Excel's own idiom -- Excel
        /// writes "Calculating (8 processors): 42%" there during a native
        /// recalculation, so a run that reports nothing reads as a hang.
        /// </summary>
        public string Describe()
        {
            var name = Phase switch
            {
                EnginePhase.Loading => "Loading workbook",
                EnginePhase.Analyzing => "Building dependency graph",
                EnginePhase.Evaluating => "Evaluating",
                EnginePhase.DataTables => "Data tables",
                _ => "Working",
            };

            if (Total == 0)
            {
                return $"LudicrousSpeed: {name}...";
            }

            var percent = (int)(Done * 100 / Total);
            return $"LudicrousSpeed: {name} {Done:N0}/{Total:N0} ({percent}%)";
        }
    }
}
