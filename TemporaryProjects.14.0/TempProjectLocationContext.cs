using System;
using System.IO;
using EnvDTE;
using Microsoft.VisualStudio.Shell;

namespace TemporaryProjects
{
    public class TempProjectLocationContext : IDisposable
    {
        readonly Property projectsLocation;
        readonly string location;

        public TempProjectLocationContext(DTE dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            projectsLocation = dte.Properties["Environment", "ProjectsAndSolution"].Item("ProjectsLocation");
            location = (string)projectsLocation.Value;
            var tempPath = string.Format(@"Temp\{0:yyyy-MM-dd}\", DateTime.Now);
            var tempLocation = Path.Combine(location, tempPath);
            projectsLocation.Value = tempLocation;
        }

        public void Dispose()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            projectsLocation.Value = location;
        }
    }
}
