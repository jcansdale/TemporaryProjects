using System;
using System.IO;
using EnvDTE;
using Microsoft.VisualStudio.Shell;

namespace TemporaryProjects
{
    public class TempProjectsLocationContext : IDisposable
    {
        readonly Property projectsLocation;
        readonly string location;

        public TempProjectsLocationContext(DTE dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            projectsLocation = dte.Properties["Environment", "ProjectsAndSolution"].Item("ProjectsLocation");
            location = (string)projectsLocation.Value;
            var tempPath = $@"Temp\{DateTime.Now:yyyy-MM-dd}\";
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
