﻿#pragma warning disable VSSDK004 // Use PackageAutoLoadFlags.None so we load ASAP!

using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;

namespace TemporaryProjects
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = false)]
    [Guid(PackageGuids.getToCodeExtenderPackageString)]
    [ProvideAutoLoad(getToCodeUIContext, PackageAutoLoadFlags.None)]
    [ProvideUIContextRule(getToCodeUIContext,
        name: "GetToCodePackageExists",
        expression: "GetToCodePackageExists",
        termNames: new[] { "GetToCodePackageExists" },
        termValues: new[]
        {
            @"ConfigSettingsStoreQuery:Packages\{D208A515-B37C-4F88-AC23-F3727FE307BD}\AllowsBackgroundLoad"
        })]
    public sealed class GetToCodeExtenderPackage : Package
    {
        const string getToCodeUIContext = "A6C01F2B-9CCB-4F06-9F82-D1835720CCFF";

        protected override void Initialize()
        {
            if (Utilities.VsVersion < 16)
                return;

            var types = new WorkflowTypes();
            if (!types.Succeeded)
                return;

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
                        "Create a new temporary project", "Bing bada bing!",
                        new DelegateCommand(_ => OnOpenTempProjectCommandExecuted(types, currentWorkflow), _ => true)
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

        private static void OnOpenTempProjectCommandExecuted(WorkflowTypes types, object currentWorkflow)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var dte = ServiceProvider.GlobalProvider.GetService(typeof(DTE)) as DTE;
            Assumes.Present(dte);
            dte.Commands.Raise(PackageGuids.guidNewTempProjectCommandPackageCmdSetString, PackageIds.NewTempProjectCommandId, null, null);

            var e = Activator.CreateInstance(types.WorkflowCompletedEventArgs, "Get To Code workflow", "After open solution");
            types.GetToCodeWorkflowViewModel_RaiseCompleted.Invoke(currentWorkflow, new object[] { e });
        }
    }
}
