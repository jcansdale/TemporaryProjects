using System;
using System.IO;
using System.ComponentModel.Design;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace TemporaryProjects
{
    internal sealed class NewTempProjectCommand
    {
        readonly DTE dte;

        private NewTempProjectCommand(IMenuCommandService commandService, DTE dte)
        {
            this.dte = dte;

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

            var projectsLocation = dte.Properties["Environment", "ProjectsAndSolution"].Item("ProjectsLocation");
            var location = (string)projectsLocation.Value;
            var tempPath = string.Format(@"Temp\{0:yyyy-MM-dd}\", DateTime.Now);
            var tempLocation = Path.Combine(location, tempPath);

            try
            {
                projectsLocation.Value = tempLocation;
                dte.ExecuteCommand("File.NewProject");
            }
            finally
            {
                projectsLocation.Value = location;
            }
        }

        public static NewTempProjectCommand Instance
        {
            get;
            private set;
        }
    }
}
