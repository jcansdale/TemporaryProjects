using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace TemporaryProjects
{
    internal sealed class StartPageExtender : IVsWindowFrameEvents
    {
        private readonly HashSet<IVsWindowFrame> trackedFrames = new HashSet<IVsWindowFrame>();
        private readonly DTE dte;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await package.JoinableTaskFactory.SwitchToMainThreadAsync();

            var dte = (DTE)await package.GetServiceAsync(typeof(DTE));
            var vsUIShell = (IVsUIShell)await package.GetServiceAsync(typeof(SVsUIShell));

            ((IVsUIShell7)vsUIShell).AdviseWindowFrameEvents(new StartPageExtender(dte, vsUIShell));
        }

        private StartPageExtender(DTE dte, IVsUIShell vsUIShell)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            this.dte = dte;

            var guid = VSConstants.StandardToolWindows.StartPage;
            if (ErrorHandler.Succeeded(vsUIShell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fFrameOnly, ref guid, out var windowFrame)) &&
                windowFrame != null)
            {
                TrackFrame(windowFrame);
            }
        }

        public void OnFrameCreated(IVsWindowFrame frame)
        {
        }

        public void OnFrameDestroyed(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            trackedFrames.Remove(frame);
        }

        public void OnFrameIsVisibleChanged(IVsWindowFrame frame, bool newIsVisible)
        {
        }

        public void OnFrameIsOnScreenChanged(IVsWindowFrame frame, bool newIsOnScreen)
        {
        }

        public void OnActiveFrameChanged(IVsWindowFrame oldFrame, IVsWindowFrame newFrame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (newFrame != null && oldFrame != newFrame)
            {
                TrackFrame(newFrame);
            }
        }

        private void TrackFrame(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (trackedFrames.Add(frame))
            {
                if (ErrorHandler.Succeeded(frame.GetGuidProperty((int)__VSFPROPID.VSFPROPID_GuidPersistenceSlot, out var guid)) &&
                    guid == VSConstants.StandardToolWindows.StartPage &&
                    ErrorHandler.Succeeded(frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out var windowPane)) &&
                    windowPane?.GetType().GetProperty(nameof(WindowPane.Content)).GetValue(windowPane) is IVsUIWpfElement wpfElement &&
                    ErrorHandler.Succeeded(wpfElement.GetFrameworkElement(out var element)) &&
                    element is FrameworkElement frameworkElement &&
                    frameworkElement.FindName("ContentHost") is Decorator contentHost)
                {
                    SetupStartPage(contentHost);
                }
            }
        }

        private void SetupStartPage(Decorator contentHost)
        {
            contentHost.LayoutUpdated += LayoutUpdated;

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

                object _ = null;
                dte.Commands.Raise(NewTempProjectCommand.CommandSet.ToString("B"), NewTempProjectCommand.CommandId, _, _);
            }
        }
    }
}
