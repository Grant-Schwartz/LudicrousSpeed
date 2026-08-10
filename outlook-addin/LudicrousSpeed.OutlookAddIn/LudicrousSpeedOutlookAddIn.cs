using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using LudicrousSpeed.OutlookAddIn.Interop;
using LudicrousSpeed.OutlookAddIn.Services;
using LudicrousSpeed.OutlookAddIn.UI;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace LudicrousSpeed.OutlookAddIn
{
    /// <summary>
    /// Stops a LudicrousSpeed-converted workbook leaving the building by
    /// accident.
    ///
    /// THE PROBLEM: Convert to Live swaps a native Excel data table for
    /// =LS.LIVE(...) cells that only resolve while the Excel add-in is loaded.
    /// That trade is invisible on the analyst's own machine -- the numbers are
    /// right there -- and it costs nothing until the file is mailed to a boss
    /// or a client, who opens it and sees a sensitivity table full of #NAME?.
    /// Nothing in Excel or Outlook connects those two moments.
    ///
    /// THE TWO PLACES IT SPEAKS:
    ///  - attach time, through <see cref="ComposeWatcher"/>, because that is
    ///    when someone still has the file in front of them and fixing it is
    ///    free;
    ///  - send time, here, because that is the only moment every path
    ///    converges on -- forwards, drag-and-drop, replies begun before this
    ///    add-in loaded -- and the only one where Outlook lets an add-in say
    ///    no.
    ///
    /// IT FAILS OPEN, ALWAYS. Every handler swallows its exceptions into
    /// <see cref="GuardLog"/>. A guard that occasionally misses a workbook is
    /// a guard people keep installed; one that occasionally refuses to send
    /// mail is uninstalled the same afternoon.
    ///
    /// NOT YET VERIFIED AGAINST LIVE OUTLOOK.
    /// </summary>
    // The GUID and ProgId are duplicated in scripts\Install-OutlookGuard.ps1,
    // which writes the COM registration by hand. Changing either here without
    // changing it there produces an add-in that installs cleanly and never
    // loads.
    [ComVisible(true)]
    [Guid("996BCB68-6439-44E2-84F0-9DF0FC563205")]
    [ProgId("LudicrousSpeed.OutlookAddIn")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class LudicrousSpeedOutlookAddIn : IDTExtensibility2
    {
        /// <summary>
        /// Escape hatch, matching the environment-variable switches the Excel
        /// add-in uses. Set LUDICROUS_OUTLOOK_GUARD=0 to load the add-in
        /// without it doing anything -- useful for deciding whether it is
        /// behind an Outlook problem without an uninstall.
        /// </summary>
        private static bool GuardEnabled =>
            Environment.GetEnvironmentVariable("LUDICROUS_OUTLOOK_GUARD") != "0";

        private readonly List<ComposeWatcher> watchers = new List<ComposeWatcher>();
        private readonly List<Outlook.Explorer> watchedExplorers = new List<Outlook.Explorer>();

        private Outlook.Application? application;

        // Held in fields rather than locals on purpose. An Outlook collection
        // is the object raising the event, and once nothing references it the
        // GC takes it and the handler simply stops firing -- minutes later,
        // silently, with no error anywhere. This is the single most common way
        // an Outlook add-in "works on my machine" and then does not.
        private Outlook.Inspectors? inspectors;
        private Outlook.Explorers? explorers;

        // Marshalling comes from the ComImport interface declaration, so the
        // parameter attributes are deliberately not repeated here.
        public void OnConnection(
            object application,
            int connectMode,
            object addInInst,
            ref Array custom)
        {
            try
            {
                this.application = application as Outlook.Application;
                if (this.application == null)
                {
                    GuardLog.Write("host is not Outlook; guard not started");
                    return;
                }

                if (!GuardEnabled)
                {
                    GuardLog.Write("guard disabled by LUDICROUS_OUTLOOK_GUARD=0");
                    return;
                }

                this.application.ItemSend += OnItemSend;

                inspectors = this.application.Inspectors;
                inspectors.NewInspector += OnNewInspector;

                explorers = this.application.Explorers;
                explorers.NewExplorer += OnNewExplorer;

                // NewExplorer only covers windows opened from now on, and
                // Outlook already has one open by the time an add-in connects.
                // Indexed rather than foreach: the collection is one-based and
                // this avoids relying on the enumerator surviving interop
                // embedding.
                for (var i = 1; i <= explorers.Count; i++)
                {
                    WatchExplorer(explorers[i]);
                }

                GuardLog.Write("guard started");
            }
            catch (Exception ex)
            {
                GuardLog.Write("OnConnection failed", ex);
            }
        }

        public void OnDisconnection(int removeMode, ref Array custom)
        {
            try
            {
                foreach (var watcher in watchers.ToList())
                {
                    watcher.Dispose();
                }

                watchers.Clear();

                foreach (var explorer in watchedExplorers)
                {
                    try
                    {
                        explorer.InlineResponse -= OnInlineResponse;
                    }
                    catch (Exception)
                    {
                    }
                }

                watchedExplorers.Clear();

                if (explorers != null)
                {
                    explorers.NewExplorer -= OnNewExplorer;
                    explorers = null;
                }

                if (inspectors != null)
                {
                    inspectors.NewInspector -= OnNewInspector;
                    inspectors = null;
                }

                if (application != null)
                {
                    application.ItemSend -= OnItemSend;
                    application = null;
                }
            }
            catch (Exception ex)
            {
                GuardLog.Write("OnDisconnection failed", ex);
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        /// <summary>
        /// The gate. Everything sendable passes through here exactly once,
        /// and setting cancel leaves the message open and unsent with the
        /// attachment still on it.
        /// </summary>
        private void OnItemSend(object item, ref bool cancel)
        {
            try
            {
                var live = AttachmentGuard.Inspect(item)
                    .Where(finding => finding.Scan.NeedsAttention)
                    .ToList();

                if (live.Count == 0)
                {
                    return;
                }

                if (GuardPrompt.ConfirmSendAnyway(live))
                {
                    GuardLog.Write($"sent anyway with {live.Count} live workbook(s) attached");
                    return;
                }

                cancel = true;
                GuardLog.Write($"send held back; {live.Count} live workbook(s) attached");
            }
            catch (Exception ex)
            {
                // Deliberately leaves cancel untouched: a crash in the guard
                // must not be able to trap someone's mail.
                GuardLog.Write("send check failed", ex);
            }
        }

        private void OnNewInspector(Outlook.Inspector inspector)
        {
            try
            {
                WatchCompose(inspector.CurrentItem);
            }
            catch (Exception ex)
            {
                GuardLog.Write("could not watch new inspector", ex);
            }
        }

        private void OnNewExplorer(Outlook.Explorer explorer)
        {
            WatchExplorer(explorer);
        }

        /// <summary>
        /// Replies written in the reading pane never open an Inspector, so
        /// without this the attach-time warning would only ever appear for
        /// people who compose in a separate window.
        /// </summary>
        private void WatchExplorer(Outlook.Explorer explorer)
        {
            try
            {
                explorer.InlineResponse += OnInlineResponse;
                watchedExplorers.Add(explorer);
            }
            catch (Exception ex)
            {
                GuardLog.Write("could not watch explorer", ex);
            }
        }

        private void OnInlineResponse(object item)
        {
            try
            {
                WatchCompose(item);
            }
            catch (Exception ex)
            {
                GuardLog.Write("could not watch inline response", ex);
            }
        }

        private void WatchCompose(object item)
        {
            if (!(item is Outlook.MailItem mail))
            {
                return;
            }

            // Sent items open in an inspector too. Watching those would hook
            // every message anybody reads, for no benefit.
            if (mail.Sent)
            {
                return;
            }

            // An inline reply popped out into its own window raises both
            // InlineResponse and NewInspector for the same underlying item,
            // which would otherwise warn twice about one attachment.
            if (watchers.Any(existing => IsSameComObject(existing.Item, mail)))
            {
                return;
            }

            watchers.Add(new ComposeWatcher(mail, watcher => watchers.Remove(watcher)));
        }

        /// <summary>
        /// Two RCWs can wrap one COM object, so reference equality says
        /// nothing. Comparing canonical IUnknown pointers is the identity test
        /// COM actually defines.
        /// </summary>
        private static bool IsSameComObject(object left, object right)
        {
            var leftUnknown = IntPtr.Zero;
            var rightUnknown = IntPtr.Zero;
            try
            {
                leftUnknown = Marshal.GetIUnknownForObject(left);
                rightUnknown = Marshal.GetIUnknownForObject(right);
                return leftUnknown == rightUnknown;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (leftUnknown != IntPtr.Zero)
                {
                    Marshal.Release(leftUnknown);
                }

                if (rightUnknown != IntPtr.Zero)
                {
                    Marshal.Release(rightUnknown);
                }
            }
        }
    }
}
