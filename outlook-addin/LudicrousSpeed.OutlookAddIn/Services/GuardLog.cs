using System;
using System.Globalization;
using System.IO;

namespace LudicrousSpeed.OutlookAddIn.Services
{
    /// <summary>
    /// Append-only trace for the things this add-in deliberately swallows.
    ///
    /// WHY IT EXISTS: the guard fails open on purpose. An attachment it cannot
    /// read, a scan that runs out of budget, an Outlook call that throws --
    /// none of those may block a send or pop a dialog, because a tool that
    /// interrupts sending mail on its own bugs gets switched off within a day.
    /// The cost of that choice is that a genuine miss looks exactly like a
    /// clean workbook, so every swallowed failure has to leave a trace
    /// somewhere. This is that somewhere.
    /// </summary>
    internal static class GuardLog
    {
        private const long MaxBytes = 1024 * 1024;

        private static readonly object Gate = new object();

        private static string LogPath =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LudicrousSpeed",
                "outlook-guard.log");

        public static void Write(string message)
        {
            try
            {
                lock (Gate)
                {
                    var path = LogPath;
                    Directory.CreateDirectory(Path.GetDirectoryName(path)!);

                    // Nobody prunes this file, so it prunes itself. Losing old
                    // lines matters far less than growing without bound in a
                    // profile directory.
                    if (File.Exists(path) && new FileInfo(path).Length > MaxBytes)
                    {
                        File.Delete(path);
                    }

                    File.AppendAllText(
                        path,
                        DateTime.Now.ToString("u", CultureInfo.InvariantCulture) + "  " + message
                            + Environment.NewLine);
                }
            }
            catch (Exception)
            {
                // A logger that throws would defeat the entire point.
            }
        }

        public static void Write(string context, Exception error)
        {
            Write(context + ": " + error.GetType().Name + " " + error.Message);
        }
    }
}
