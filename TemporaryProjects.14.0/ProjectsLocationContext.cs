using System;
using EnvDTE;
using Microsoft.VisualStudio.Shell;

namespace TemporaryProjects
{
    public class ProjectsLocationEvent
    {
        readonly DTE dte;
        readonly CommandEvents commandEvents;
        IDisposable context;

        public ProjectsLocationEvent(DTE dte, Guid guid, int id)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.dte = dte;
            commandEvents = dte.Events.CommandEvents[guid.ToString("B"), id];
            commandEvents.BeforeExecute += BeforeExecute;
            commandEvents.AfterExecute += AfterExecute;
        }

        void BeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            context = new TempProjectLocationContext(dte);
        }

        void AfterExecute(string Guid, int ID, object CustomIn, object CustomOut)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            context?.Dispose();
            context = null;
        }
    }
}
