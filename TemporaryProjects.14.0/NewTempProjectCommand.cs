using System;
using System.ComponentModel.Design;
using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace TemporaryProjects
{
    internal sealed class NewTempProjectCommand
    {
        readonly DTE dte;


        NewTempProjectCommand(IMenuCommandService commandService, DTE dte) : this(dte)
        {
            var menuCommandID = new CommandID(PackageGuids.guidNewTempProjectCommandPackageCmdSet, PackageIds.NewTempProjectCommandId);
            var menuItem = new MenuCommand((s, e) => Execute(), menuCommandID);
            commandService.AddCommand(menuItem);
        }

        NewTempProjectCommand() : this(ServiceProvider.GlobalProvider.GetService(typeof(DTE)) as DTE)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
        }


        NewTempProjectCommand(DTE dte)
        {
            Assumes.Present(dte);
            this.dte = dte;
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
        void Execute()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var saveNewProjects = dte.Properties["Environment", "ProjectsAndSolution"].Item("SaveNewProjects");
            var oldValue = saveNewProjects.Value;
            try
            {
                using (new TempProjectLocationContext(dte))
                {
                    saveNewProjects.Value = false;
                    dte.ExecuteCommand("File.NewProject");
                }
            }
            finally
            {
                saveNewProjects.Value = oldValue;
            }
        }

        public static NewTempProjectCommand Instance
        {
            get;
            private set;
        }
    }
}
