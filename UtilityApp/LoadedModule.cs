using System.Windows;
using UtilityApp.Contracts;

namespace UtilityApp
{
    internal sealed class LoadedModule
    {
        public string AssemblyPath { get; set; }

        public string ModuleTypeName { get; set; }

        public IUtilityModule Instance { get; set; }

        public FrameworkElement CachedView { get; set; }

        public string Id
        {
            get { return Instance == null ? string.Empty : Instance.Id; }
        }

        public string Name
        {
            get { return Instance == null ? string.Empty : Instance.Name; }
        }

        public string Description
        {
            get { return Instance == null ? string.Empty : Instance.Description; }
        }

        public bool EnabledByDefault
        {
            get { return Instance != null && Instance.EnabledByDefault; }
        }
    }
}
