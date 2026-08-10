using System;
using System.Collections.Generic;
using System.Linq;
using LudicrousSpeed.OutlookAddIn.UI;
using Outlook = Microsoft.Office.Interop.Outlook;
using WinFormsTimer = System.Windows.Forms.Timer;

namespace LudicrousSpeed.OutlookAddIn.Services
{
    /// <summary>
    /// Watches one message being composed and speaks up the moment a live
    /// workbook is attached to it, rather than waiting for Send.
    ///
    /// WHY THE CHECK IS DEFERRED INSTEAD OF RUN INSIDE AttachmentAdd: the
    /// event fires while Outlook is still writing the attachment, so asking
    /// for its bytes right then can fail outright, and showing a dialog from
    /// inside the event re-enters Outlook mid-operation. A short timer solves
    /// both -- the attachment has finished landing, Outlook has finished its
    /// own work, and selecting five files in the attach dialog collapses into
    /// one prompt instead of five, because each add restarts the same timer.
    ///
    /// This is the courtesy warning. It is allowed to miss things -- an
    /// attachment carried in by a forward or a drag-and-drop onto a message
    /// list never raises AttachmentAdd on an item we are watching. The
    /// ItemSend gate is what actually guarantees nothing gets out unnoticed.
    /// </summary>
    internal sealed class ComposeWatcher : IDisposable
    {
        private const int SettleMilliseconds = 500;

        private readonly Outlook.MailItem item;
        private readonly Action<ComposeWatcher> onFinished;
        private readonly WinFormsTimer timer;

        /// <summary>
        /// Files already mentioned for this message. Without it, attaching a
        /// second workbook would re-announce the first one, and the warning
        /// starts reading as noise.
        /// </summary>
        private readonly HashSet<string> announced =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private bool disposed;

        public ComposeWatcher(Outlook.MailItem item, Action<ComposeWatcher> onFinished)
        {
            this.item = item;
            this.onFinished = onFinished;

            timer = new WinFormsTimer { Interval = SettleMilliseconds };
            timer.Tick += OnSettled;

            item.AttachmentAdd += OnAttachmentAdd;
            item.Unload += OnUnload;
        }

        public object Item => item;

        private void OnAttachmentAdd(Outlook.Attachment attachment)
        {
            if (disposed)
            {
                return;
            }

            // Restarting rather than starting is what coalesces a multi-file
            // attach into a single prompt.
            timer.Stop();
            timer.Start();
        }

        private void OnSettled(object? sender, EventArgs e)
        {
            timer.Stop();
            if (disposed)
            {
                return;
            }

            try
            {
                var fresh = AttachmentGuard.Inspect(item)
                    .Where(finding => finding.Scan.NeedsAttention)
                    .Where(finding => announced.Add(finding.FileName))
                    .ToList();

                if (fresh.Count > 0)
                {
                    GuardPrompt.NoticeOnAttach(fresh);
                }
            }
            catch (Exception ex)
            {
                // Same rule as everywhere else here: the compose window keeps
                // working, ItemSend still gets its turn.
                GuardLog.Write("attach-time check failed", ex);
            }
        }

        private void OnUnload()
        {
            Dispose();
        }

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }

            disposed = true;

            try
            {
                timer.Tick -= OnSettled;
                timer.Dispose();
            }
            catch (Exception)
            {
            }

            // By the time Unload has fired the item is already gone, so
            // detaching from it is expected to throw and is done anyway for
            // the paths that dispose earlier.
            try
            {
                item.AttachmentAdd -= OnAttachmentAdd;
                item.Unload -= OnUnload;
            }
            catch (Exception)
            {
            }

            onFinished(this);
        }
    }
}
