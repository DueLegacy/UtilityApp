using UtilityApp.Contracts;

namespace UtilityApp.Modules.Templates
{
    public class TemplatesModule : IUtilityModule
    {
        public string Id
        {
            get { return "templates"; }
        }

        public string Name
        {
            get { return "Templates"; }
        }

        public string Description
        {
            get { return "Store reusable text templates and copy them to the clipboard."; }
        }

        public bool EnabledByDefault
        {
            get { return true; }
        }

        public System.Windows.FrameworkElement CreateView(IHostContext hostContext)
        {
            return new TemplatesModuleView(hostContext);
        }
    }
}
