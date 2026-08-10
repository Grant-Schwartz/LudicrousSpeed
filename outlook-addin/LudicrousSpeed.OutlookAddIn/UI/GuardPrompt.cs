using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using LudicrousSpeed.OutlookAddIn.Services;

namespace LudicrousSpeed.OutlookAddIn.UI
{
    /// <summary>
    /// The two things the guard ever says to anyone.
    ///
    /// WHY PLAIN MESSAGE BOXES: the warning has to be readable in the second
    /// somebody is reaching for Send, and it has to name the fix. A custom
    /// form would add a designer, a layout to keep working across DPI
    /// settings, and nothing a user would notice. The Excel add-in reports
    /// through MessageBox for the same reason.
    ///
    /// WHY THE SEND PROMPT DEFAULTS TO "No": the expensive mistake is sending
    /// a broken model to a client, and the cheap one is spending ten seconds
    /// re-attaching a file. Enter should land on the cheap one.
    /// </summary>
    internal static class GuardPrompt
    {
        private const string Caption = "LudicrousSpeed";

        /// <summary>
        /// Named sheets stop being useful once there are more of them than
        /// someone will read standing up.
        /// </summary>
        private const int MaxSheetsNamed = 4;

        /// <returns>True to let the send proceed.</returns>
        public static bool ConfirmSendAnyway(IReadOnlyList<AttachmentFinding> findings)
        {
            var message = new StringBuilder();
            message.AppendLine(
                findings.Count == 1
                    ? "This message has an attachment with LudicrousSpeed live cells still in it:"
                    : "This message has attachments with LudicrousSpeed live cells still in them:");
            message.AppendLine();

            foreach (var finding in findings)
            {
                message.AppendLine("    " + Describe(finding));
            }

            message.AppendLine();
            message.AppendLine(
                "Anyone without LudicrousSpeed installed sees #NAME? in those cells "
                + "instead of numbers.");
            message.AppendLine();
            message.AppendLine(RestoreInstructions());

            if (findings.Any(f => f.Scan.UsesLegacyName))
            {
                message.AppendLine();
                message.AppendLine(
                    "Some of those cells use the retired WS.LIVE name, which no longer "
                    + "resolves even here -- re-run Convert to Live if you need them working.");
            }

            if (findings.Any(f => f.Scan.HasConversionMetadata))
            {
                message.AppendLine();
                message.AppendLine(
                    "The workbook also carries a hidden _LudicrousSpeed_DataTables sheet. "
                    + "Restore Native leaves it behind on purpose so a failed restore stays "
                    + "recoverable; delete it by hand if the file is going outside your firm.");
            }

            message.AppendLine();
            message.Append("Send anyway?");

            var answer = MessageBox.Show(
                message.ToString(),
                Caption,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            return answer == DialogResult.Yes;
        }

        public static void NoticeOnAttach(IReadOnlyList<AttachmentFinding> findings)
        {
            var message = new StringBuilder();
            foreach (var finding in findings)
            {
                message.AppendLine(Describe(finding));
            }

            message.AppendLine();
            message.AppendLine(
                "Anyone without LudicrousSpeed installed sees #NAME? in those cells "
                + "instead of numbers.");
            message.AppendLine();
            message.AppendLine(RestoreInstructions());
            message.AppendLine();
            message.Append("You will get one more chance to turn back at Send.");

            MessageBox.Show(
                message.ToString(),
                Caption,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private static string RestoreInstructions()
        {
            return "To fix it: open the workbook in Excel, click Restore Native on the "
                + "LudicrousSpeed tab, save, then attach it again.";
        }

        private static string Describe(AttachmentFinding finding)
        {
            var cells = finding.Scan.LiveCellCount;
            var summary = string.Format(
                CultureInfo.CurrentCulture,
                "{0} -- {1:N0} live {2}",
                finding.FileName,
                cells,
                cells == 1 ? "cell" : "cells");

            var sheets = finding.Scan.Sheets;
            if (sheets.Count == 0)
            {
                return summary;
            }

            var named = string.Join(", ", sheets.Take(MaxSheetsNamed));
            if (sheets.Count > MaxSheetsNamed)
            {
                named += $", +{sheets.Count - MaxSheetsNamed} more";
            }

            return summary + " on " + named;
        }
    }
}
