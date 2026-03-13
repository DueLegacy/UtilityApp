using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using UtilityApp.Contracts;

namespace UtilityApp.Modules.Templates
{
    public partial class TemplatesModuleView : UserControl
    {
        private readonly IHostContext _hostContext;
        private readonly List<TemplateItem> _templates = new List<TemplateItem>();
        private readonly string _templateStorePath;

        private string _selectedTemplateId;
        private bool _isEditMode;
        private bool _suppressTemplateSelectionChanged;

        public TemplatesModuleView(IHostContext hostContext)
        {
            InitializeComponent();

            _hostContext = hostContext;
            _templateStorePath = Path.Combine(
                _hostContext == null ? AppDomain.CurrentDomain.BaseDirectory : _hostContext.ConfigDirectoryPath,
                "templates.xml");

            LoadTemplates();
            RefreshTemplateList(null);
            SetEditMode(false, false);
        }

        private void TemplateSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshTemplateList(_selectedTemplateId);
        }

        private void TemplateListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressTemplateSelectionChanged)
            {
                return;
            }

            var selectedItem = TemplateListBox.SelectedItem as ListBoxItem;
            if (selectedItem == null)
            {
                if (!_isEditMode)
                {
                    ClearEditor();
                    SetEditMode(false, false);
                }

                return;
            }

            var templateId = selectedItem.Tag as string;
            var template = FindTemplateById(templateId);
            if (template == null)
            {
                return;
            }

            LoadTemplateIntoEditor(template);
            CopyTemplateToClipboard(template.Name, template.Content);
        }

        private void EditTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_selectedTemplateId))
            {
                Log("Edit mode blocked: no template selected.");
                return;
            }

            SetEditMode(true, true);
            Log("Enabled template edit mode.");
        }

        private void DeleteTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var currentTemplate = FindTemplateById(_selectedTemplateId);
            if (currentTemplate == null)
            {
                Log("Template delete blocked: no template selected.");
                return;
            }

            if (MessageBox.Show(
                    string.Format("Delete the template '{0}'?", currentTemplate.Name),
                    "Delete Template",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning,
                    MessageBoxResult.No) != MessageBoxResult.Yes)
            {
                return;
            }

            if (MessageBox.Show(
                    string.Format("Delete '{0}' permanently? This cannot be undone.", currentTemplate.Name),
                    "Confirm Delete",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning,
                    MessageBoxResult.No) != MessageBoxResult.Yes)
            {
                return;
            }

            var removedIndex = _templates.IndexOf(currentTemplate);
            if (removedIndex < 0)
            {
                Log("Template delete failed: selected template is unavailable.");
                return;
            }

            _templates.RemoveAt(removedIndex);
            if (!SaveTemplates())
            {
                _templates.Insert(removedIndex, currentTemplate);
                SortTemplates();
                RefreshTemplateList(currentTemplate.Id);
                return;
            }

            _selectedTemplateId = null;
            RefreshTemplateList(null);
            Log(string.Format("Deleted template: {0}.", currentTemplate.Name));
        }

        private void NewTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            _selectedTemplateId = null;

            _suppressTemplateSelectionChanged = true;
            try
            {
                TemplateListBox.SelectedIndex = -1;
            }
            finally
            {
                _suppressTemplateSelectionChanged = false;
            }

            ClearEditor();
            SetEditMode(true, true);
            Log("Started a new template draft.");
        }

        private void SaveTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var templateName = (TemplateNameTextBox.Text ?? string.Empty).Trim();
            var templateContent = TemplateContentTextBox.Text ?? string.Empty;

            if (string.IsNullOrWhiteSpace(templateName))
            {
                Log("Template save blocked: name is empty.");
                return;
            }

            var currentTemplate = FindTemplateById(_selectedTemplateId);
            if (HasTemplateNameConflict(templateName, currentTemplate == null ? null : currentTemplate.Id))
            {
                Log(string.Format("Template save blocked: duplicate name '{0}'.", templateName));
                return;
            }

            var isNewTemplate = currentTemplate == null;
            if (isNewTemplate)
            {
                currentTemplate = new TemplateItem
                {
                    Id = Guid.NewGuid().ToString("N")
                };
                _templates.Add(currentTemplate);
            }

            currentTemplate.Name = templateName;
            currentTemplate.Content = templateContent;

            SortTemplates();

            if (!SaveTemplates())
            {
                return;
            }

            _selectedTemplateId = currentTemplate.Id;
            RefreshTemplateList(currentTemplate.Id);
            LoadTemplateIntoEditor(currentTemplate);
            SetEditMode(false, false);
            Log(isNewTemplate
                ? string.Format("Created template: {0}.", templateName)
                : string.Format("Updated template: {0}.", templateName));
        }

        private void CopyTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var templateName = (TemplateNameTextBox.Text ?? string.Empty).Trim();
            var templateContent = TemplateContentTextBox.Text ?? string.Empty;

            CopyTemplateToClipboard(templateName, templateContent);
        }

        private void LoadTemplates()
        {
            _templates.Clear();
            _selectedTemplateId = null;

            if (!File.Exists(_templateStorePath))
            {
                return;
            }

            try
            {
                var serializer = new XmlSerializer(typeof(TemplateStore));
                using (var stream = File.OpenRead(_templateStorePath))
                {
                    var templateStore = serializer.Deserialize(stream) as TemplateStore;
                    if (templateStore != null && templateStore.Templates != null)
                    {
                        foreach (var template in templateStore.Templates)
                        {
                            if (template == null || string.IsNullOrWhiteSpace(template.Name))
                            {
                                continue;
                            }

                            if (string.IsNullOrWhiteSpace(template.Id))
                            {
                                template.Id = Guid.NewGuid().ToString("N");
                            }

                            if (template.Content == null)
                            {
                                template.Content = string.Empty;
                            }

                            _templates.Add(template);
                        }
                    }
                }

                SortTemplates();
                Log(string.Format("Loaded {0} templates from {1}.", _templates.Count, ToAppRelativePath(_templateStorePath)));
            }
            catch (Exception ex)
            {
                Log(string.Format("Template load failed: {0}", ex.Message));
            }
        }

        private bool SaveTemplates()
        {
            try
            {
                var directoryPath = Path.GetDirectoryName(_templateStorePath);
                if (!string.IsNullOrWhiteSpace(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                var serializer = new XmlSerializer(typeof(TemplateStore));
                using (var stream = File.Create(_templateStorePath))
                {
                    serializer.Serialize(stream, new TemplateStore(_templates));
                }

                return true;
            }
            catch (Exception ex)
            {
                Log(string.Format("Template save failed: {0}", ex.Message));
                return false;
            }
        }

        private void RefreshTemplateList(string templateIdToSelect)
        {
            var searchTerm = (TemplateSearchTextBox.Text ?? string.Empty).Trim();
            var visibleTemplates = new List<TemplateItem>();

            foreach (var template in _templates)
            {
                if (TemplateMatchesSearch(template, searchTerm))
                {
                    visibleTemplates.Add(template);
                }
            }

            _suppressTemplateSelectionChanged = true;
            try
            {
                TemplateListBox.Items.Clear();

                foreach (var template in visibleTemplates)
                {
                    var listItem = new ListBoxItem
                    {
                        Content = template.Name,
                        Tag = template.Id,
                        Style = (Style)FindResource("TemplateListItemStyle")
                    };

                    TemplateListBox.Items.Add(listItem);

                    if (string.Equals(template.Id, templateIdToSelect, StringComparison.OrdinalIgnoreCase))
                    {
                        TemplateListBox.SelectedItem = listItem;
                    }
                }
            }
            finally
            {
                _suppressTemplateSelectionChanged = false;
            }

            TemplateListCountText.Text = visibleTemplates.Count.ToString();

            if (!string.IsNullOrWhiteSpace(templateIdToSelect))
            {
                var selectedTemplate = FindTemplateById(templateIdToSelect);
                if (selectedTemplate != null)
                {
                    LoadTemplateIntoEditor(selectedTemplate);
                }
            }
            else if (!_isEditMode || visibleTemplates.Count == 0)
            {
                ClearEditor();
                SetEditMode(false, false);
            }
        }

        private static bool TemplateMatchesSearch(TemplateItem template, string searchTerm)
        {
            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                return true;
            }

            return !string.IsNullOrWhiteSpace(template.Name)
                && template.Name.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private TemplateItem FindTemplateById(string templateId)
        {
            if (string.IsNullOrWhiteSpace(templateId))
            {
                return null;
            }

            foreach (var template in _templates)
            {
                if (string.Equals(template.Id, templateId, StringComparison.OrdinalIgnoreCase))
                {
                    return template;
                }
            }

            return null;
        }

        private bool HasTemplateNameConflict(string templateName, string currentTemplateId)
        {
            foreach (var template in _templates)
            {
                if (string.Equals(template.Id, currentTemplateId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (string.Equals(template.Name, templateName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private void LoadTemplateIntoEditor(TemplateItem template)
        {
            if (template == null)
            {
                return;
            }

            _selectedTemplateId = template.Id;
            TemplateNameTextBox.Text = template.Name ?? string.Empty;
            TemplateContentTextBox.Text = template.Content ?? string.Empty;
            SetEditMode(false, false);
        }

        private void ClearEditor()
        {
            TemplateNameTextBox.Text = string.Empty;
            TemplateContentTextBox.Text = string.Empty;
        }

        private void SetEditMode(bool isEditMode, bool focusEditor)
        {
            _isEditMode = isEditMode;

            TemplateNameTextBox.IsReadOnly = !_isEditMode;
            TemplateContentTextBox.IsReadOnly = !_isEditMode;
            EditTemplateButton.IsEnabled = !_isEditMode && !string.IsNullOrWhiteSpace(_selectedTemplateId);
            DeleteTemplateButton.IsEnabled = !_isEditMode && !string.IsNullOrWhiteSpace(_selectedTemplateId);
            SaveTemplateButton.IsEnabled = _isEditMode;

            if (!focusEditor)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(_selectedTemplateId))
            {
                TemplateNameTextBox.Focus();
                return;
            }

            TemplateContentTextBox.Focus();
        }

        private void SortTemplates()
        {
            _templates.Sort(delegate(TemplateItem left, TemplateItem right)
            {
                var leftName = left == null ? string.Empty : left.Name ?? string.Empty;
                var rightName = right == null ? string.Empty : right.Name ?? string.Empty;
                return StringComparer.OrdinalIgnoreCase.Compare(leftName, rightName);
            });
        }

        private bool CopyTemplateToClipboard(string templateName, string templateContent)
        {
            if (string.IsNullOrEmpty(templateContent))
            {
                Log("Clipboard copy skipped: template content is empty.");
                return false;
            }

            try
            {
                Clipboard.SetText(templateContent);
                Log(string.IsNullOrWhiteSpace(templateName)
                    ? "Copied template content to clipboard."
                    : string.Format("Copied template content: {0}.", templateName));
                return true;
            }
            catch (Exception ex)
            {
                Log(string.Format("Clipboard copy failed: {0}", ex.Message));
                return false;
            }
        }

        private string ToAppRelativePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            var basePath = (_hostContext == null ? AppDomain.CurrentDomain.BaseDirectory : _hostContext.ApplicationRootPath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            if (path.StartsWith(basePath, StringComparison.OrdinalIgnoreCase))
            {
                var relativePath = path.Substring(basePath.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                return string.IsNullOrWhiteSpace(relativePath) ? ".\\" : string.Format(".\\{0}", relativePath);
            }

            return path;
        }

        private void Log(string message)
        {
            if (_hostContext != null)
            {
                _hostContext.Log(message);
            }
        }
    }
}

