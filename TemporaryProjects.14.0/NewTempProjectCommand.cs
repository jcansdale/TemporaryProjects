using System;
using System.ComponentModel.Design;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace TemporaryProjects
{
    internal sealed class NewTempProjectCommand
    {
        private NewTempProjectCommand(IMenuCommandService commandService, DTE dte)
        {
            var menuCommandID = new CommandID(PackageGuids.guidNewTempProjectCommandPackageCmdSet, PackageIds.NewTempProjectCommandId);
            var menuItem = new MenuCommand((s, e) => Execute(dte), menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await package.JoinableTaskFactory.SwitchToMainThreadAsync();

            var commandService = (IMenuCommandService)await package.GetServiceAsync((typeof(IMenuCommandService)));
            var dte = (DTE)await package.GetServiceAsync((typeof(DTE)));
            Instance = new NewTempProjectCommand(commandService, dte);
        }

        // You can test this command in the current VS instance like this:
        // https://github.com/jcansdale/TestDriven.Net-Issues/wiki/Test-With...VS-SDK
        [STAThread]
        static void Execute(DTE dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            using (new TempProjectsLocationContext(dte))
            {
                dte.ExecuteCommand("File.NewProject");
            }
        }

        public static NewTempProjectCommand Instance
        {
            get;
            private set;
        }
    }
}
