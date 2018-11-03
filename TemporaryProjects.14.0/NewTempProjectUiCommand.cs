using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Windows.Input;

namespace TemporaryProjects
{
    internal sealed class NewTempProjectUiCommand : ICommand
    {
        private DTE dte;

        public event EventHandler CanExecuteChanged { add { } remove { } }

        public bool CanExecute(object parameter) => true;

        public void Execute(object parameter)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (dte == null)
                dte = (DTE)ServiceProvider.GlobalProvider.GetService(typeof(DTE));

            dte.Commands.Raise(PackageGuids.guidNewTempProjectCommandPackageCmdSetString, PackageIds.NewTempProjectCommandId, null, null);
        }
    }
}
