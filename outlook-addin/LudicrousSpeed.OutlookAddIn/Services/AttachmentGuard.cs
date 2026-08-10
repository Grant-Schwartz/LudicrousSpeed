using System;
using System.Collections.Generic;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace LudicrousSpeed.OutlookAddIn.Services
{
    internal sealed class AttachmentFinding
    {
        public AttachmentFinding(string fileName, WorkbookScan scan)
        {
            FileName = fileName;
            Scan = scan;
        }

        public string FileName { get; }

        public WorkbookScan Scan { get; }
    }

    /// <summary>
    /// Turns an Outlook item into the list of its attachments that still hold
    /// LudicrousSpeed live cells.
    ///
    /// WHY IT WRITES A TEMP FILE: an attachment is not a file on disk until
    /// you ask for one. Outlook's object model exposes the bytes only through
    /// Attachment.SaveAsFile, so the scan gets a real path to open as a zip
    /// and the copy is deleted straight afterwards, whatever happens.
    ///
    /// WHY IT FILTERS ON EXTENSION FIRST: without that, every PDF, image and
    /// signature logo on every outgoing mail would be written to disk and read
    /// back before being rejected. The extension check keeps the common case
    /// -- a message with no workbook attached -- at zero I/O.
    /// </summary>
    internal static class AttachmentGuard
    {
        private static string TempFolder =>
            Path.Combine(Path.GetTempPath(), "LudicrousSpeedGuard");

        public static List<AttachmentFinding> Inspect(object item)
        {
            var findings = new List<AttachmentFinding>();
            Outlook.Attachments? attachments = TryGetAttachments(item);
            if (attachments == null)
            {
                return findings;
            }

            int count;
            try
            {
                count = attachments.Count;
            }
            catch (Exception ex)
            {
                GuardLog.Write("could not count attachments", ex);
                return findings;
            }

            // One-based, and by index rather than foreach: enumerating an
            // Outlook collection while Outlook is mid-send has a habit of
            // throwing, and an index walk lets one bad entry be skipped
            // without losing the rest.
            for (var i = 1; i <= count; i++)
            {
                try
                {
                    var finding = InspectOne(attachments[i]);
                    if (finding != null)
                    {
                        findings.Add(finding);
                    }
                }
                catch (Exception ex)
                {
                    GuardLog.Write($"attachment {i} could not be inspected", ex);
                }
            }

            return findings;
        }

        private static AttachmentFinding? InspectOne(Outlook.Attachment attachment)
        {
            var fileName = attachment.FileName;
            if (!LiveCellScanner.LooksLikeWorkbook(fileName))
            {
                return null;
            }

            Directory.CreateDirectory(TempFolder);
            var temp = Path.Combine(
                TempFolder,
                Guid.NewGuid().ToString("N") + Path.GetExtension(fileName));

            try
            {
                // Throws for link-style attachments (a OneDrive or SharePoint
                // reference has no bytes to hand over). Nothing to scan in
                // that case, and the log line is the only honest outcome.
                attachment.SaveAsFile(temp);

                var scan = LiveCellScanner.Scan(temp);
                if (scan.Inconclusive != null)
                {
                    // Deliberately not a reason to discard the finding. A scan
                    // that ran out of budget after seeing live cells has
                    // already proved its point; throwing that away would turn
                    // "too big to finish reading" into "looks fine", which is
                    // the one answer this must never give by accident. The log
                    // line is for the other case -- found nothing and did not
                    // finish -- which is indistinguishable from clean.
                    GuardLog.Write($"'{fileName}' not fully checked: {scan.Inconclusive}");
                }

                return new AttachmentFinding(fileName, scan);
            }
            finally
            {
                TryDelete(temp);
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex)
            {
                // A leftover copy of someone's model in %TEMP% is worth a line
                // in the log even though it cannot be helped here.
                GuardLog.Write($"could not delete temporary copy {path}", ex);
            }
        }

        /// <summary>
        /// ItemSend fires for mail, meeting requests, task requests and
        /// anything else sendable, and those types share no common interface
        /// that exposes Attachments. Asking late-bound covers all of them
        /// without a cast per type, and the items that genuinely have no
        /// attachments simply fail the call.
        /// </summary>
        private static Outlook.Attachments? TryGetAttachments(object item)
        {
            try
            {
                dynamic sendable = item;
                return sendable.Attachments as Outlook.Attachments;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
