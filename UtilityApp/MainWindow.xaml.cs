using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using UtilityApp.Contracts;

namespace UtilityApp
{
    public partial class MainWindow : Window
    {
        private const string OverviewNavigationId = "overview";

        private readonly Dictionary<string, bool> _moduleEnabledStates = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ListBoxItem> _moduleItems = new Dictionary<string, ListBoxItem>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, LoadedModule> _modulesById = new Dictionary<string, LoadedModule>(StringComparer.OrdinalIgnoreCase);
        private readonly List<LoadedModule> _modules = new List<LoadedModule>();
        private readonly string _modulesRootPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Modules");
        private readonly HostContext _hostContext;

        private bool _suppressModuleSelectionChanged;

        public MainWindow()
        {
            InitializeComponent();

            _hostContext = new HostContext(
                AppDomain.CurrentDomain.BaseDirectory,
                AppendActivity);

            RefreshModuleCatalog(OverviewNavigationId, false);
            AppendActivity("Host UI shell initialized.");
        }

        private void ModuleList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressModuleSelectionChanged)
            {
                return;
            }

            var navigationId = GetSelectedNavigationId();
            if (string.IsNullOrWhiteSpace(navigationId))
            {
                return;
            }

            ApplyNavigationSelection(navigationId);
        }

        private void OpenOperationsButton_Click(object sender, RoutedEventArgs e)
        {
            ShowOperationsMenu();
            AppendActivity("Opened consolidated Operations menu.");
        }

        private void HomeButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.Equals(GetSelectedNavigationId(), OverviewNavigationId, StringComparison.OrdinalIgnoreCase))
            {
                ShowMainSurface();
                ShowOverviewPanel();
                AppendActivity("Opened Home workspace.");
                return;
            }

            SelectNavigationItem(OverviewNavigationId);
        }

        private void DiscoverModulesButton_Click(object sender, RoutedEventArgs e)
        {
            var currentNavigationId = GetSelectedNavigationId();
            RefreshModuleCatalog(currentNavigationId, true);
        }

        private void ValidateConfigButton_Click(object sender, RoutedEventArgs e)
        {
            ShowOperationsMenu();
            AppendActivity(string.Format("Validated host configuration view against {0} discovered module DLL(s).", _modules.Count));
        }

        private void OpenLogButton_Click(object sender, RoutedEventArgs e)
        {
            ShowOperationsMenu();
            AppendActivity("Opened in-app log preview.");

            if (ActivityLogList.Items.Count > 0)
            {
                ActivityLogList.ScrollIntoView(ActivityLogList.Items[0]);
            }
        }

        private void RefreshModuleCatalog(string navigationIdToSelect, bool logDiscovery)
        {
            var previousStates = new Dictionary<string, bool>(_moduleEnabledStates, StringComparer.OrdinalIgnoreCase);

            DiscoverModules(previousStates, logDiscovery);
            BuildNavigation(navigationIdToSelect);
            InitializeModuleControls();
            UpdateModuleAvailability();
            UpdateOperationsSummary();

            var selectedNavigationId = GetSelectedNavigationId();
            if (string.IsNullOrWhiteSpace(selectedNavigationId))
            {
                selectedNavigationId = OverviewNavigationId;
            }

            ApplyNavigationSelection(selectedNavigationId);
        }

        private void DiscoverModules(Dictionary<string, bool> previousStates, bool logDiscovery)
        {
            _modules.Clear();
            _modulesById.Clear();
            _moduleEnabledStates.Clear();

            if (!Directory.Exists(_modulesRootPath))
            {
                if (logDiscovery)
                {
                    AppendActivity("Discovery scan complete: .\\Modules directory is missing.");
                }

                return;
            }

            foreach (var assemblyPath in Directory.GetFiles(_modulesRootPath, "*.dll", SearchOption.TopDirectoryOnly))
            {
                DiscoverModuleAssembly(assemblyPath, previousStates);
            }

            _modules.Sort(delegate(LoadedModule left, LoadedModule right)
            {
                var leftName = left == null ? string.Empty : left.Name ?? string.Empty;
                var rightName = right == null ? string.Empty : right.Name ?? string.Empty;
                return StringComparer.OrdinalIgnoreCase.Compare(leftName, rightName);
            });

            if (logDiscovery)
            {
                AppendActivity(string.Format("Discovery scan complete: {0} module DLL(s) found in .\\Modules.", _modules.Count));
            }
        }

        private void DiscoverModuleAssembly(string assemblyPath, Dictionary<string, bool> previousStates)
        {
            Assembly assembly;
            try
            {
                assembly = Assembly.LoadFrom(assemblyPath);
            }
            catch (Exception ex)
            {
                AppendActivity(string.Format("Module assembly load failed: {0} ({1})", ToAppRelativePath(assemblyPath), ex.Message));
                return;
            }

            foreach (var type in GetLoadableTypes(assembly, assemblyPath))
            {
                if (type == null || !typeof(IUtilityModule).IsAssignableFrom(type) || !type.IsClass || type.IsAbstract)
                {
                    continue;
                }

                IUtilityModule instance;
                try
                {
                    instance = (IUtilityModule)Activator.CreateInstance(type);
                }
                catch (Exception ex)
                {
                    AppendActivity(string.Format(
                        "Module instantiation failed: {0} ({1})",
                        type.FullName,
                        ex.Message));
                    continue;
                }

                var moduleId = (instance.Id ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(moduleId))
                {
                    AppendActivity(string.Format("Skipped module with missing id from {0}.", ToAppRelativePath(assemblyPath)));
                    continue;
                }

                if (_modulesById.ContainsKey(moduleId))
                {
                    AppendActivity(string.Format("Skipped duplicate module id: {0}.", moduleId));
                    continue;
                }

                var loadedModule = new LoadedModule
                {
                    AssemblyPath = assemblyPath,
                    ModuleTypeName = type.FullName,
                    Instance = instance
                };

                _modules.Add(loadedModule);
                _modulesById[moduleId] = loadedModule;
                _moduleEnabledStates[moduleId] = previousStates.ContainsKey(moduleId)
                    ? previousStates[moduleId]
                    : loadedModule.EnabledByDefault;
            }
        }

        private IEnumerable<Type> GetLoadableTypes(Assembly assembly, string assemblyPath)
        {
            try
            {
                return assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                foreach (var loaderException in ex.LoaderExceptions)
                {
                    if (loaderException == null)
                    {
                        continue;
                    }

                    AppendActivity(string.Format(
                        "Module type inspection failed: {0} ({1})",
                        ToAppRelativePath(assemblyPath),
                        loaderException.Message));
                }

                return ex.Types;
            }
        }

        private void BuildNavigation(string navigationIdToSelect)
        {
            _suppressModuleSelectionChanged = true;
            try
            {
                ModuleList.Items.Clear();
                _moduleItems.Clear();

                ModuleList.Items.Add(CreateNavigationItem(OverviewNavigationId, "Overview", true));

                foreach (var module in _modules)
                {
                    var navigationItem = CreateNavigationItem(module.Id, module.Name, IsModuleEnabled(module.Id));
                    ModuleList.Items.Add(navigationItem);
                    _moduleItems[module.Id] = navigationItem;
                }

                if (!SelectNavigationItemInternal(navigationIdToSelect))
                {
                    SelectNavigationItemInternal(OverviewNavigationId);
                }
            }
            finally
            {
                _suppressModuleSelectionChanged = false;
            }
        }

        private ListBoxItem CreateNavigationItem(string navigationId, string title, bool isEnabled)
        {
            return new ListBoxItem
            {
                Content = title,
                Tag = navigationId,
                IsEnabled = isEnabled,
                Style = (Style)FindResource("ModuleListItemStyle")
            };
        }

        private bool SelectNavigationItem(string navigationId)
        {
            if (!SelectNavigationItemInternal(navigationId))
            {
                return false;
            }

            ApplyNavigationSelection(GetSelectedNavigationId());
            return true;
        }

        private bool SelectNavigationItemInternal(string navigationId)
        {
            for (int index = 0; index < ModuleList.Items.Count; index++)
            {
                var item = ModuleList.Items[index] as ListBoxItem;
                if (item == null)
                {
                    continue;
                }

                var itemNavigationId = item.Tag as string;
                if (string.Equals(itemNavigationId, navigationId, StringComparison.OrdinalIgnoreCase))
                {
                    ModuleList.SelectedIndex = index;
                    return true;
                }
            }

            return false;
        }

        private string GetSelectedNavigationId()
        {
            var selectedItem = ModuleList.SelectedItem as ListBoxItem;
            return selectedItem == null ? null : selectedItem.Tag as string;
        }

        private void ApplyNavigationSelection(string navigationId)
        {
            if (string.IsNullOrWhiteSpace(navigationId))
            {
                return;
            }

            if (string.Equals(navigationId, OverviewNavigationId, StringComparison.OrdinalIgnoreCase))
            {
                ActiveModuleTitle.Text = "Overview";
                ShowMainSurface();
                ShowOverviewPanel();
                AppendActivity("Selected module view: Overview.");
                return;
            }

            LoadedModule module;
            if (!_modulesById.TryGetValue(navigationId, out module))
            {
                SelectNavigationItem(OverviewNavigationId);
                return;
            }

            if (!IsModuleEnabled(module.Id))
            {
                SelectNavigationItem(OverviewNavigationId);
                return;
            }

            ActiveModuleTitle.Text = module.Name;
            ShowMainSurface();

            var moduleView = EnsureModuleView(module);
            if (moduleView != null)
            {
                ShowModuleView(moduleView);
                AppendActivity(string.Format("Selected module view: {0}.", module.Name));
                return;
            }

            ShowGenericModuleWorkspace(module);
            AppendActivity(string.Format("Selected module view: {0}.", module.Name));
        }

        private FrameworkElement EnsureModuleView(LoadedModule module)
        {
            if (module == null)
            {
                return null;
            }

            if (module.CachedView != null)
            {
                return module.CachedView;
            }

            try
            {
                module.CachedView = module.Instance.CreateView(_hostContext);
                return module.CachedView;
            }
            catch (Exception ex)
            {
                AppendActivity(string.Format("Module view creation failed: {0} ({1})", module.Name, ex.Message));
                return null;
            }
        }

        private void InitializeModuleControls()
        {
            ModuleTogglePanel.Children.Clear();

            foreach (var module in _modules)
            {
                ModuleTogglePanel.Children.Add(CreateModuleControlRow(module));
            }
        }

        private Border CreateModuleControlRow(LoadedModule module)
        {
            var rowBorder = new Border
            {
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(14, 12, 14, 12),
                Background = (Brush)FindResource("ShellBackgroundBrush"),
                BorderBrush = (Brush)FindResource("PanelBorderBrush"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8)
            };

            var rowGrid = new Grid();
            rowGrid.ColumnDefinitions.Add(new ColumnDefinition());
            rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var detailsPanel = new StackPanel();
            detailsPanel.Children.Add(new TextBlock
            {
                Text = module.Name,
                FontFamily = new FontFamily("Bahnschrift"),
                FontSize = 15,
                FontWeight = FontWeights.SemiBold,
                Foreground = (Brush)FindResource("HeadingBrush")
            });
            detailsPanel.Children.Add(new TextBlock
            {
                Margin = new Thickness(0, 4, 0, 0),
                Text = module.Description,
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                Foreground = (Brush)FindResource("BodyBrush")
            });
            detailsPanel.Children.Add(new TextBlock
            {
                Margin = new Thickness(0, 6, 0, 0),
                Text = ToAppRelativePath(module.AssemblyPath),
                FontFamily = new FontFamily("Consolas"),
                FontSize = 12,
                Foreground = (Brush)FindResource("BodyBrush")
            });

            var enabledCheckBox = new CheckBox
            {
                Tag = module.Id,
                Content = "Enabled",
                IsChecked = IsModuleEnabled(module.Id),
                VerticalAlignment = VerticalAlignment.Center,
                FontFamily = new FontFamily("Bahnschrift"),
                FontSize = 13,
                Foreground = (Brush)FindResource("HeadingBrush")
            };
            enabledCheckBox.Checked += ModuleEnabledCheckBox_Changed;
            enabledCheckBox.Unchecked += ModuleEnabledCheckBox_Changed;

            Grid.SetColumn(enabledCheckBox, 1);

            rowGrid.Children.Add(detailsPanel);
            rowGrid.Children.Add(enabledCheckBox);

            rowBorder.Child = rowGrid;
            return rowBorder;
        }

        private void ModuleEnabledCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            var enabledCheckBox = sender as CheckBox;
            var moduleId = enabledCheckBox == null ? null : enabledCheckBox.Tag as string;
            if (string.IsNullOrWhiteSpace(moduleId))
            {
                return;
            }

            _moduleEnabledStates[moduleId] = enabledCheckBox.IsChecked == true;
            UpdateModuleAvailability();

            LoadedModule module;
            _modulesById.TryGetValue(moduleId, out module);
            var moduleName = module == null ? moduleId : module.Name;

            if (string.Equals(GetSelectedNavigationId(), moduleId, StringComparison.OrdinalIgnoreCase) && !IsModuleEnabled(moduleId))
            {
                SelectNavigationItem(OverviewNavigationId);
                AppendActivity(string.Format("Disabled module: {0}. Returned to Overview.", moduleName));
                return;
            }
            AppendActivity(string.Format(
                "Module {0}: {1}.",
                moduleName,
                IsModuleEnabled(moduleId) ? "enabled" : "disabled"));
        }

        private void UpdateModuleAvailability()
        {
            foreach (var moduleItemEntry in _moduleItems)
            {
                moduleItemEntry.Value.IsEnabled = IsModuleEnabled(moduleItemEntry.Key);
            }

            UpdateOverviewSummary();
            UpdateOperationsSummary();
        }

        private void UpdateOverviewSummary()
        {
            var availableModuleCount = GetAvailableModuleCount();
            var enabledModuleCount = GetEnabledModuleCount();

            AvailableModulesCountText.Text = availableModuleCount.ToString();
            EnabledModulesCountText.Text = enabledModuleCount.ToString();
            DisabledModulesCountText.Text = (availableModuleCount - enabledModuleCount).ToString();
        }

        private void UpdateOperationsSummary()
        {
            DiscoveryStateText.Text = _modules.Count == 0
                ? "No Module DLLs Found"
                : string.Format("{0} Module DLL{1} Ready", _modules.Count, _modules.Count == 1 ? string.Empty : "s");
            DiscoveryStateText.Foreground = (Brush)FindResource(_modules.Count == 0 ? "WarnBrush" : "GoodBrush");
        }

        private int GetAvailableModuleCount()
        {
            return _modules.Count;
        }

        private int GetEnabledModuleCount()
        {
            var enabledModuleCount = 0;

            foreach (var moduleState in _moduleEnabledStates)
            {
                if (moduleState.Value)
                {
                    enabledModuleCount++;
                }
            }

            return enabledModuleCount;
        }

        private bool IsModuleEnabled(string moduleId)
        {
            bool isEnabled;
            return _moduleEnabledStates.TryGetValue(moduleId, out isEnabled) && isEnabled;
        }

        private void ShowOverviewPanel()
        {
            OverviewPanel.Visibility = Visibility.Visible;
            ModuleViewHost.Visibility = Visibility.Collapsed;
            ModuleWorkspaceCard.Visibility = Visibility.Collapsed;
        }

        private void ShowModuleView(FrameworkElement moduleView)
        {
            OverviewPanel.Visibility = Visibility.Collapsed;
            ModuleWorkspaceCard.Visibility = Visibility.Collapsed;
            ModuleViewHost.Content = moduleView;
            ModuleViewHost.Visibility = Visibility.Visible;
        }

        private void ShowGenericModuleWorkspace(LoadedModule module)
        {
            OverviewPanel.Visibility = Visibility.Collapsed;
            ModuleViewHost.Visibility = Visibility.Collapsed;
            ModuleWorkspaceCard.Visibility = Visibility.Visible;
            SelectedModuleStateText.Text = IsModuleEnabled(module.Id) ? "Enabled" : "Disabled";
            SelectedModuleAssemblyText.Text = ToAppRelativePath(module.AssemblyPath);
            SelectedModuleDescriptionText.Text = string.IsNullOrWhiteSpace(module.Description)
                ? "No module description available."
                : module.Description;
        }

        private string ToAppRelativePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            var basePath = AppDomain.CurrentDomain.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (path.StartsWith(basePath, StringComparison.OrdinalIgnoreCase))
            {
                var relativePath = path.Substring(basePath.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                return string.IsNullOrWhiteSpace(relativePath) ? ".\\" : string.Format(".\\{0}", relativePath);
            }

            return path;
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
            AppendActivityToFile(entry);
            ActivityLogList.Items.Insert(0, entry);

            while (ActivityLogList.Items.Count > 40)
            {
                ActivityLogList.Items.RemoveAt(ActivityLogList.Items.Count - 1);
            }
        }

        private void AppendActivityToFile(string entry)
        {
            try
            {
                var logsDirectoryPath = _hostContext == null
                    ? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs")
                    : _hostContext.LogsDirectoryPath;
                var logFilePath = Path.Combine(logsDirectoryPath, DateTime.Now.ToString("yyyyMMdd") + ".log");

                File.AppendAllText(logFilePath, entry + Environment.NewLine);
            }
            catch
            {
                // Logging must never crash the host.
            }
        }
    }
}
