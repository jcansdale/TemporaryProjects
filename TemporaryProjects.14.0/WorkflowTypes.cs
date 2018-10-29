using Microsoft.VisualStudio.Imaging.Interop;
using System;
using System.Reflection;
using System.Windows.Input;

namespace TemporaryProjects
{
    internal sealed class WorkflowTypes
    {
        public bool Succeeded { get; }

        public Type WorkflowHostView { get; }
        public Type WorkflowHostViewModel { get; }
        public Type GetToCodeWorkflowViewModel { get; }
        public Type GetToCodeAction { get; }
        public Type WorkflowCompletedEventArgs { get; }

        public PropertyInfo WorkflowHostView_Instance { get; }
        public PropertyInfo WorkflowHostViewModel_CurrentWorkflow { get; }
        public EventInfo WorkflowHostViewModel_PropertyChanged;
        public PropertyInfo GetToCodeWorkflowViewModel_Actions;
        public MethodInfo GetToCodeWorkflowViewModel_NotifyPropertyChanged;
        public MethodInfo GetToCodeWorkflowViewModel_RaiseCompleted;
        public ConstructorInfo GetToCodeAction_Constructor;

        public WorkflowTypes()
        {
            var assembly = Assembly.Load("Microsoft.VisualStudio.Shell.UI.Internal");

            if ((WorkflowHostView = assembly.GetType("Microsoft.VisualStudio.PlatformUI.GetToCode.WorkflowHostView")) != null &&
                (WorkflowHostViewModel = assembly.GetType("Microsoft.VisualStudio.PlatformUI.GetToCode.WorkflowHostViewModel")) != null &&
                (GetToCodeWorkflowViewModel = assembly.GetType("Microsoft.VisualStudio.PlatformUI.GetToCode.GetToCodeWorkflowViewModel")) != null &&
                (GetToCodeAction = assembly.GetType("Microsoft.VisualStudio.PlatformUI.GetToCode.GetToCodeAction")) != null &&
                (WorkflowCompletedEventArgs = assembly.GetType("Microsoft.VisualStudio.PlatformUI.GetToCode.WorkflowCompletedEventArgs")) != null &&
                (WorkflowHostView_Instance = WorkflowHostView.GetProperty("Instance", BindingFlags.NonPublic | BindingFlags.Static)) != null &&
                (WorkflowHostViewModel_CurrentWorkflow = WorkflowHostViewModel.GetProperty("CurrentWorkflow")) != null &&
                (WorkflowHostViewModel_PropertyChanged = WorkflowHostViewModel.GetEvent("PropertyChanged")) != null &&
                (GetToCodeWorkflowViewModel_Actions = GetToCodeWorkflowViewModel.GetProperty("Actions")) != null &&
                (GetToCodeWorkflowViewModel_NotifyPropertyChanged = GetToCodeWorkflowViewModel.GetMethod("NotifyPropertyChanged", BindingFlags.NonPublic | BindingFlags.Instance)) != null &&
                (GetToCodeWorkflowViewModel_RaiseCompleted = GetToCodeWorkflowViewModel.GetMethod("RaiseCompleted", BindingFlags.NonPublic | BindingFlags.Instance)) != null &&
                (GetToCodeAction_Constructor = GetToCodeAction.GetConstructor(
                    BindingFlags.NonPublic | BindingFlags.Instance,
                    null,
                    new[] { typeof(ImageMoniker), typeof(string), typeof(string), typeof(ICommand) },
                    new ParameterModifier[4])) != null)
            {
                Succeeded = true;
            }
        }
    }
}
