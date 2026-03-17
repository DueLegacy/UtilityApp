using UtilityApp.Contracts;

namespace UtilityApp.Modules.BusinessRequirements
{
    public class BusinessRequirementsModule : IUtilityModule
    {
        public string Id
        {
            get { return "business-requirements"; }
        }

        public string Name
        {
            get { return "Business Requirements"; }
        }

        public string Description
        {
            get { return "Business Requirement summaries."; }
        }

        public bool EnabledByDefault
        {
            get { return true; }
        }

        public System.Windows.FrameworkElement CreateView(IHostContext hostContext)
        {
            return new BusinessRequirementsModuleView(hostContext);
        }
    }
}
