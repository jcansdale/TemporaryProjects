using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace TemporaryProjects
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(StartPageToolWindowString, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class StartPageExtenderPackage : AsyncPackage
    {
        public const string PackageGuidString = "3dd793ab-452e-4e50-9eaf-57ad755aff1d";
        public const string StartPageToolWindowString = "387cb18d-6153-4156-9257-9ac3f9207bbe";

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await StartPageExtender.InitializeAsync(this);
        }
    }
}
