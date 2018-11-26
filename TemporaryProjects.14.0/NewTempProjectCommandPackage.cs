using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;

namespace TemporaryProjects
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuids.newTempProjectCommandPackageString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class NewTempProjectCommandPackage : AsyncPackage
    {
        ProjectsLocationEvent saveAllContext;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var dte = await GetServiceAsync(typeof(SDTE)) as DTE;
            saveAllContext = new ProjectsLocationEvent(dte,
                VSConstants.GUID_VSStandardCommandSet97,
                (int)VSConstants.VSStd97CmdID.SaveSolution /*File.SaveAll*/);

            await NewTempProjectCommand.InitializeAsync(this);
        }
    }
}
