using System.Windows;

namespace UtilityApp.Contracts
{
    public interface IUtilityModule
    {
        string Id { get; }

        string Name { get; }

        string Description { get; }

        bool EnabledByDefault { get; }

        FrameworkElement CreateView(IHostContext hostContext);
    }
}
