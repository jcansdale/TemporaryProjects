using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace TemporaryProjects
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(PackageGuids.startPageExtenderPackageString)]
    [ProvideAutoLoad(PackageGuids.startPageToolWindowString, PackageAutoLoadFlags.BackgroundLoad)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class StartPageExtenderPackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            if (Utilities.VsVersion < 15)
            {
                // Doesn't work with Visual Studio 2015
                return;
            }

            await JoinableTaskFactory.SwitchToMainThreadAsync();
            var vsUIShell = (IVsUIShell7)await GetServiceAsync(typeof(SVsUIShell));
            StartPageExtender.Initialize(vsUIShell);
        }
    }
}
