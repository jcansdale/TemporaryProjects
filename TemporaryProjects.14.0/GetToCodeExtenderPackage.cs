using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using Task = System.Threading.Tasks.Task;

namespace TemporaryProjects
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(PackageGuids.getToCodeExtenderPackageString)]
    [ProvideAutoLoad(UIContextGuids.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class GetToCodeExtenderPackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            if (Utilities.VsVersion < 16)
                return;

            var types = new WorkflowTypes();
            if (!types.Succeeded)
                return;

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var window = (FrameworkElement)types.WorkflowHostView_Instance.GetValue(null);
            window.Height += 32;

            var workflowHostViewModel = window.DataContext;
            types.WorkflowHostViewModel_PropertyChanged.AddEventHandler(workflowHostViewModel, new PropertyChangedEventHandler(WorkflowHostViewModel_PropertyChanged));

            UpdateCurrentWorkflow();

            void WorkflowHostViewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
            {
                if (e.PropertyName == types.WorkflowHostViewModel_CurrentWorkflow.Name)
                    UpdateCurrentWorkflow();
            }

            void UpdateCurrentWorkflow()
            {
                var currentWorkflow = types.WorkflowHostViewModel_CurrentWorkflow.GetValue(workflowHostViewModel);
                if (currentWorkflow.GetType() == types.GetToCodeWorkflowViewModel)
                {
                    var newAction = types.GetToCodeAction_Constructor.Invoke(new object[]
                    {
                        KnownMonikers.NewTestGroup,
                        "Create a new temporary project",
                        "",
                        new NewTempProjectUiCommand()
                    });

                    var currentActions = (object[])types.GetToCodeWorkflowViewModel_Actions.GetValue(currentWorkflow);
                    var newActions = (object[])Array.CreateInstance(types.GetToCodeAction, currentActions.Length + 1);
                    currentActions.CopyTo(newActions, 0);
                    newActions[currentActions.Length] = newAction;

                    types.GetToCodeWorkflowViewModel_Actions.SetValue(currentWorkflow, newActions);
                    types.GetToCodeWorkflowViewModel_NotifyPropertyChanged.Invoke(currentWorkflow, new object[] { types.GetToCodeWorkflowViewModel_Actions.Name });
                }
            }
        }
    }
}
