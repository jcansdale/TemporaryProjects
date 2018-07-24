using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using Task = System.Threading.Tasks.Task;

namespace TemporaryProjects
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(PackageGuids.startPageExtenderPackage14String)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(PackageGuids.startPageToolWindowString, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class StartPageExtenderPackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync();

            var dte = (DTE)await GetServiceAsync(typeof(DTE));
            var shell = (IVsShell7)await GetServiceAsync(typeof(SVsShell));
            switch (dte.Version)
            {
                case "15.0":
                    // This package only works in Visual Studio 2017
                    await shell.LoadPackageAsync(PackageGuids.startPageExtenderPackage15);
                    break;
            }
        }
    }
}
