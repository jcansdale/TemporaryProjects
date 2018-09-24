using System;
using System.Windows;
using System.Windows.Controls;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace TemporaryProjects
{
    internal sealed class StartPageExtender : IVsWindowFrameEvents
    {
        private readonly DTE dte;
        private const string CookieName = "TemporaryProjects.StartPageExtender.Cookie";

        // You can test this in the current Visual Studio instance like this:
        // https://github.com/jcansdale/TestDriven.Net-Issues/wiki/Test-With...VS-SDK
        [STAThread]
        public static void Initialize(IVsUIShell7 vsUIShell, DTE dte)
        {
            ReadviseWindowFrameEvents(vsUIShell, dte);
        }

        private static void ReadviseWindowFrameEvents(IVsUIShell7 vsUIShell, DTE dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (Cookie is uint cookie)
            {
                // Ensure that we only have one event sink installed for the app domain if initialize is called
                // multiple times while testing.
                vsUIShell.UnadviseWindowFrameEvents(cookie);
            }

            Cookie = vsUIShell.AdviseWindowFrameEvents(new StartPageExtender(dte, (IVsUIShell)vsUIShell));
        }

        // Store cookie in app domain wide variable because assembly might be loaded multiple times while testing.
        static uint? Cookie
        {
            get => (uint?)AppDomain.CurrentDomain.GetData(CookieName);
            set => AppDomain.CurrentDomain.SetData(CookieName, value);
        }

        private StartPageExtender(DTE dte, IVsUIShell vsUIShell)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            this.dte = dte;

            if (ErrorHandler.Succeeded(vsUIShell.FindToolWindow(
                    (uint)__VSFINDTOOLWIN.FTW_fFrameOnly, VSConstants.StandardToolWindows.StartPage, out var windowFrame)) &&
                windowFrame != null)
            {
                TrackFrame(windowFrame);
            }
        }

        public void OnFrameCreated(IVsWindowFrame frame)
        {
            // When the frame for the start page is created, it is empty and has no content. Not only that, but
            // the content of the frame seems to be recreated every time the tab is activated, so we have to
            // extend its content in OnFrameIsOnScreenChanged. This may happen multiple times with the same frame. If we only
            // did this once, our button would be gone if the user switched to another tab and then back to the start page.
        }

        public void OnFrameDestroyed(IVsWindowFrame frame)
        {
        }

        public void OnFrameIsVisibleChanged(IVsWindowFrame frame, bool newIsVisible)
        {
        }

        public void OnFrameIsOnScreenChanged(IVsWindowFrame frame, bool newIsOnScreen)
        {
            // If you try to make any improvements here, please make sure not only that the button is visible
            // when the start page is opened, but also:
            // 1. When the user switches to another tab and back to the start page
            // 2. When the start page is undocked by dragging the tab
            // 3. When the start page is docked again (this currently *doesn't* work)
            // 4. The button is never duplicated, for example when switching focus to a different tool window and back

            // I've tried many approaches such as doing this in OnFrameIsVisibleChanged, OnActiveFrameChanged,
            // IVsWindowFrameNotify.OnShow using IVsWindowFrame2.Advise and even DTE.Events.WindowEvents.
            // With all of these approaches there are cases where the button is not recreated because the event is not fired,
            // and in some cases it's even fired twice. The current solution is a compromise - the common case works and
            // undocking the start page is presumably not that common.

            ThreadHelper.ThrowIfNotOnUIThread();

            if (newIsOnScreen)
            {
                TrackFrame(frame);
            }
        }

        public void OnActiveFrameChanged(IVsWindowFrame oldFrame, IVsWindowFrame newFrame)
        {
        }

        private void TrackFrame(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (ErrorHandler.Succeeded(frame.GetGuidProperty((int)__VSFPROPID.VSFPROPID_GuidPersistenceSlot, out var guid)) &&
                guid == VSConstants.StandardToolWindows.StartPage &&
                ErrorHandler.Succeeded(frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out var windowPane)) &&
                // The type of windowPane is Microsoft.VisualStudio.Shell.ToolWindowPane (which inherits from WindowPane),
                // but if we tried to cast it, the cast would fail at runtime in VS 2017 where the real type comes from
                // Microsoft.VisualStudio.Shell.15.0, whereas we're referencing
                // Microsoft.VisualStudio.Shell.14.0.
                // Both can't be referenced at the same time so we have to use dynamic as a workaround.
                (windowPane as dynamic).Content is IVsUIWpfElement wpfElement &&
                ErrorHandler.Succeeded(wpfElement.GetFrameworkElement(out var element)) &&
                element is FrameworkElement frameworkElement &&
                frameworkElement.FindName("ContentHost") is Decorator contentHost)
            {
                SetupStartPage(contentHost);
            }
        }

        private void SetupStartPage(Decorator contentHost)
        {
            // The start page loads in phases and there will usually be about 2 layout updates before our target UI is available,
            // but it's also possible that the UI is fully loaded already, so we make a call to LayoutUpdated immediately.
            // This will especially happen if we found it using FindToolWindow in the constructor.

            contentHost.LayoutUpdated += LayoutUpdated;
            LayoutUpdated(null, null);

            void LayoutUpdated(object sender, EventArgs e)
            {
                var recentProjectsPanel = FindRecentProjectsPanel(contentHost);
                if (recentProjectsPanel != null)
                {
                    AddNewTemporaryProjectButton(recentProjectsPanel);
                    contentHost.LayoutUpdated -= LayoutUpdated;
                }
            }
        }

        private static Grid FindRecentProjectsPanel(Decorator contentHost)
        {
            if (contentHost.Child is ContentControl start &&
                start.Template.FindName("NewProjectsListView", start) is ContentControl newProjectsListView &&
                newProjectsListView.Template.FindName("RecentProjectsPanel", newProjectsListView) is Grid recentProjectsPanel)
            {
                return recentProjectsPanel;
            }

            return null;
        }

        private void AddNewTemporaryProjectButton(Grid recentProjectsPanel)
        {
            var button = new NewTemporaryProjectButton(dte)
            {
                Style = ((Button)recentProjectsPanel.FindName("MoreTemplatesButton")).Style
            };

            recentProjectsPanel.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
            Grid.SetRow(button, recentProjectsPanel.RowDefinitions.Count - 1);

            recentProjectsPanel.Children.Add(button);
        }

        private sealed class NewTemporaryProjectButton : Button
        {
            private readonly DTE dte;

            public NewTemporaryProjectButton(DTE dte)
            {
                this.dte = dte;
            }

            public override void OnApplyTemplate()
            {
                base.OnApplyTemplate();

                var textBlock = (TextBlock)GetTemplateChild("MoreTemplatesText");
                textBlock.Text = "Create new temporary project...";
            }

            protected override void OnClick()
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                base.OnClick();

                dte.Commands.Raise(PackageGuids.guidNewTempProjectCommandPackageCmdSetString, PackageIds.NewTempProjectCommandId, null, null);
            }
        }
    }
}
