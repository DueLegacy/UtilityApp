using Microsoft.Win32;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using UtilityApp.Contracts;

namespace UtilityApp.Modules.BusinessRequirements
{
    public partial class BusinessRequirementsModuleView : UserControl
    {
        private readonly BusinessRequirementAnalyzer _analyzer = new BusinessRequirementAnalyzer();
        private readonly IHostContext _hostContext;

        public BusinessRequirementsModuleView(IHostContext hostContext)
        {
            InitializeComponent();

            _hostContext = hostContext;
            YearTextBox.Text = DateTime.Today.Year.ToString(CultureInfo.InvariantCulture);
            PeriodTypeComboBox.SelectedIndex = 0;
            UpdatePeriodValueChoices();
            Loaded += BusinessRequirementsModuleView_Loaded;
            ResetOutputState();
        }

        private void BusinessRequirementsModuleView_Loaded(object sender, RoutedEventArgs e)
        {
            if (PeriodValueComboBox.Items.Count == 0 || (!string.Equals(GetSelectedPeriodType(), "a", StringComparison.OrdinalIgnoreCase) && PeriodValueComboBox.SelectedIndex < 0))
            {
                UpdatePeriodValueChoices();
            }
        }

        private void WorkbookPathTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateOutputPreview();
        }

        private void BrowseWorkbookButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                Title = "Select Workbook"
            };

            if (File.Exists(WorkbookPathTextBox.Text))
            {
                dialog.FileName = WorkbookPathTextBox.Text;
            }

            if (dialog.ShowDialog() == true)
            {
                WorkbookPathTextBox.Text = dialog.FileName;
            }
        }


        private void PeriodTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded)
            {
                return;
            }

            UpdatePeriodValueChoices();
        }

        private void RunAnalysisButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var options = BuildAnalysisOptions();
                var result = _analyzer.Analyze(options);
                var outputFilePath = GetOutputFilePath(result.WorkbookPath);

                File.WriteAllText(outputFilePath, result.FormattedReportText ?? string.Empty, new UTF8Encoding(false));
                ApplyOutputResult(result, outputFilePath);

                Log(string.Format(
                    "Business requirements analysis completed for {0} ({1}). Saved {2}.",
                    ToAppRelativePath(result.WorkbookPath),
                    result.PeriodLabel,
                    ToAppRelativePath(outputFilePath)));


            }
            catch (Exception ex)
            {
                LastRunSummaryTextBlock.Text = "Run failed: " + ex.Message;
                Log("Business requirements analysis failed: " + ex.Message);
            }
        }

        private BusinessRequirementAnalysisOptions BuildAnalysisOptions()
        {
            int year;
            if (!int.TryParse((YearTextBox.Text ?? string.Empty).Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out year)
                || year < 1
                || year > 9999)
            {
                throw new InvalidOperationException("Year is invalid.");
            }

            var periodType = GetSelectedPeriodType();
            if (string.IsNullOrWhiteSpace(periodType))
            {
                throw new InvalidOperationException("Period type is required.");
            }

            int? periodValue = null;
            if (!string.Equals(periodType, "a", StringComparison.OrdinalIgnoreCase))
            {
                var selectedPeriodValue = PeriodValueComboBox.SelectedItem as ComboBoxItem;
                int parsedPeriodValue;
                if (selectedPeriodValue == null || !int.TryParse(selectedPeriodValue.Tag as string, NumberStyles.None, CultureInfo.InvariantCulture, out parsedPeriodValue))
                {
                    throw new InvalidOperationException("Period value is required.");
                }

                periodValue = parsedPeriodValue;
            }

            return new BusinessRequirementAnalysisOptions
            {
                WorkbookPath = (WorkbookPathTextBox.Text ?? string.Empty).Trim(),
                Year = year,
                PeriodType = periodType,
                PeriodValue = periodValue,
                IncludeProgressDetails = ShowProgressCheckBox.IsChecked == true
            };
        }

        private string GetSelectedPeriodType()
        {
            var selectedItem = PeriodTypeComboBox.SelectedItem as ComboBoxItem;
            return selectedItem == null ? null : selectedItem.Tag as string;
        }

        private void UpdatePeriodValueChoices()
        {
            var selectedPeriodType = GetSelectedPeriodType();
            PeriodValueComboBox.Items.Clear();

            if (string.Equals(selectedPeriodType, "a", StringComparison.OrdinalIgnoreCase))
            {
                PeriodValueComboBox.IsEnabled = false;
                PeriodValueComboBox.SelectedIndex = -1;
                return;
            }

            PeriodValueComboBox.IsEnabled = true;

            var maxValue = 12;
            if (string.Equals(selectedPeriodType, "q", StringComparison.OrdinalIgnoreCase))
            {
                maxValue = 4;
            }
            else if (string.Equals(selectedPeriodType, "s", StringComparison.OrdinalIgnoreCase))
            {
                maxValue = 2;
            }

            for (var value = 1; value <= maxValue; value++)
            {
                PeriodValueComboBox.Items.Add(new ComboBoxItem
                {
                    Content = value.ToString(CultureInfo.InvariantCulture),
                    Tag = value.ToString(CultureInfo.InvariantCulture)
                });
            }

            var defaultValue = 1;
            if (string.Equals(selectedPeriodType, "m", StringComparison.OrdinalIgnoreCase))
            {
                defaultValue = DateTime.Today.Month;
            }
            else if (string.Equals(selectedPeriodType, "q", StringComparison.OrdinalIgnoreCase))
            {
                defaultValue = ((DateTime.Today.Month - 1) / 3) + 1;
            }
            else if (string.Equals(selectedPeriodType, "s", StringComparison.OrdinalIgnoreCase))
            {
                defaultValue = DateTime.Today.Month <= 6 ? 1 : 2;
            }

            if (defaultValue > maxValue)
            {
                defaultValue = maxValue;
            }

            PeriodValueComboBox.SelectedIndex = defaultValue - 1;
        }

        private void ApplyOutputResult(BusinessRequirementAnalysisResult result, string outputFilePath)
        {
            OutputFilePathTextBlock.Text = ToAppRelativePath(outputFilePath);
            LastRunSummaryTextBlock.Text = string.Format(
                CultureInfo.InvariantCulture,
                "Period: {0}. Total records: {1}. After filter: {2}.",
                result.PeriodLabel,
                result.TotalRecords,
                result.FilteredRecords);
        }

        private void ResetOutputState()
        {
            UpdateOutputPreview();
            LastRunSummaryTextBlock.Text = "No report generated yet.";
        }

        private void UpdateOutputPreview()
        {
            OutputFilePathTextBlock.Text = ToAppRelativePath(GetOutputFilePath(WorkbookPathTextBox.Text));
        }

        private string GetOutputFilePath(string workbookPath)
        {
            var basePath = GetApplicationRootPath();
            var normalizedWorkbookPath = (workbookPath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedWorkbookPath))
            {
                return Path.Combine(basePath, "stat.txt");
            }

            string fullWorkbookPath;
            try
            {
                fullWorkbookPath = Path.IsPathRooted(normalizedWorkbookPath)
                    ? Path.GetFullPath(normalizedWorkbookPath)
                    : Path.GetFullPath(Path.Combine(basePath, normalizedWorkbookPath));
            }
            catch
            {
                return Path.Combine(basePath, "stat.txt");
            }

            var directoryPath = Path.GetDirectoryName(fullWorkbookPath);
            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                directoryPath = basePath;
            }

            return Path.Combine(directoryPath, "stat.txt");
        }


        private string GetApplicationRootPath()
        {
            return _hostContext == null ? AppDomain.CurrentDomain.BaseDirectory : _hostContext.ApplicationRootPath;
        }

        private string ToAppRelativePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            var basePath = GetApplicationRootPath().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

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



