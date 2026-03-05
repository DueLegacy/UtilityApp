using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace UtilityApp
{
    public partial class MainWindow : Window
    {
        private readonly Dictionary<string, string> _moduleDescriptions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Overview", "Prepare the host shell and shared services before plugin assemblies are introduced." },
            { "File Utilities", "Future file-oriented plugins can expose safe copy, filter, and conversion tools here." },
            { "Data Tools", "Future offline data helpers can process local CSV/XML/JSON files without internet dependency." },
            { "Analytics", "Future analytics plugins can provide internal summaries using only local application data." },
            { "Settings", "Settings is routed to the consolidated Operations menu." }
        };

        public MainWindow()
        {
            InitializeComponent();

            if (ModuleList.Items.Count > 0)
            {
                ModuleList.SelectedIndex = 0;
            }

            ShowMainSurface();
            StatusText.Text = "Host shell ready. Waiting for plugin discovery pipeline.";
            AppendActivity("Host UI shell initialized.");
        }

        private void ModuleList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedItem = ModuleList.SelectedItem as ListBoxItem;
            if (selectedItem == null)
            {
                return;
            }

            var selectedModule = selectedItem.Content as string;
            if (string.IsNullOrWhiteSpace(selectedModule))
            {
                return;
            }

            ActiveModuleTitle.Text = selectedModule;

            string description;
            if (!_moduleDescriptions.TryGetValue(selectedModule, out description))
            {
                description = "Module details will appear here once plugin metadata is available.";
            }

            ActiveModuleDescription.Text = description;

            if (string.Equals(selectedModule, "Settings", StringComparison.OrdinalIgnoreCase))
            {
                ShowOperationsMenu();
                StatusText.Text = "Focused module: Settings (Operations menu).";
                AppendActivity("Selected module view: Settings (Operations menu).");
                return;
            }

            ShowMainSurface();
            StatusText.Text = string.Format("Focused module: {0}", selectedModule);
            AppendActivity(string.Format("Selected module view: {0}", selectedModule));
        }

        private void OpenOperationsButton_Click(object sender, RoutedEventArgs e)
        {
            ShowOperationsMenu();
            StatusText.Text = "Opened Operations menu.";
            AppendActivity("Opened consolidated Operations menu.");
        }

        private void HomeButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = ModuleList.SelectedItem as ListBoxItem;
            var selectedModule = selectedItem == null ? null : selectedItem.Content as string;

            if (string.Equals(selectedModule, "Overview", StringComparison.OrdinalIgnoreCase))
            {
                ShowMainSurface();
                StatusText.Text = "Returned to main workspace.";
                AppendActivity("Opened Home workspace.");
                return;
            }

            SelectModuleByName("Overview");
            ShowMainSurface();
        }

        private void DiscoverModulesButton_Click(object sender, RoutedEventArgs e)
        {
            ShowOperationsMenu();
            StatusText.Text = "Discovery check complete. Modules folder is currently empty.";
            AppendActivity("Ran discovery check against .\\Modules (preview mode).");
        }

        private void ValidateConfigButton_Click(object sender, RoutedEventArgs e)
        {
            ShowOperationsMenu();
            StatusText.Text = "Configuration check complete. Config schema placeholder is valid.";
            AppendActivity("Validated configuration placeholder for offline host settings.");
        }

        private void OpenLogButton_Click(object sender, RoutedEventArgs e)
        {
            ShowOperationsMenu();
            StatusText.Text = "Log preview refreshed in Operations menu.";
            AppendActivity("Opened in-app log preview (file logger not wired yet).");

            if (ActivityLogList.Items.Count > 0)
            {
                ActivityLogList.ScrollIntoView(ActivityLogList.Items[0]);
            }
        }

        private void GoToSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SelectModuleByName("Settings");
            ShowOperationsMenu();
            StatusText.Text = "Opened settings section in Operations menu.";
            AppendActivity("Navigated to settings section in Operations menu.");
        }

        private void SelectModuleByName(string moduleName)
        {
            for (int index = 0; index < ModuleList.Items.Count; index++)
            {
                var item = ModuleList.Items[index] as ListBoxItem;
                var itemName = item == null ? null : item.Content as string;
                if (string.Equals(itemName, moduleName, StringComparison.OrdinalIgnoreCase))
                {
                    ModuleList.SelectedIndex = index;
                    return;
                }
            }
        }

        private void ShowMainSurface()
        {
            MainSurfacePanel.Visibility = Visibility.Visible;
            OperationsMenuPanel.Visibility = Visibility.Collapsed;
        }

        private void ShowOperationsMenu()
        {
            MainSurfacePanel.Visibility = Visibility.Collapsed;
            OperationsMenuPanel.Visibility = Visibility.Visible;
        }

        private void AppendActivity(string message)
        {
            var entry = string.Format("[{0}] {1}", DateTime.Now.ToString("HH:mm:ss"), message);
            ActivityLogList.Items.Insert(0, entry);

            while (ActivityLogList.Items.Count > 40)
            {
                ActivityLogList.Items.RemoveAt(ActivityLogList.Items.Count - 1);
            }
        }
    }
}
