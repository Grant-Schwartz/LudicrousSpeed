using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Newtonsoft.Json;
using LudicrousSpeed.ExcelAddIn.Models;

namespace LudicrousSpeed.ExcelAddIn.Interop
{
    internal sealed class NativeEngineClient
    {
        private const string WindowsDll = "ludicrous_engine.dll";

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
                return EngineResponse.Failed(
                    $"Could not find {WindowsDll}. Build the Rust engine and copy the native library beside the Excel add-in.");
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

        [DllImport(WindowsDll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern IntPtr ludicrous_run_json(string requestJson);

        [DllImport(WindowsDll, CallingConvention = CallingConvention.Cdecl)]
        private static extern void ludicrous_free_string(IntPtr value);
    }
}
