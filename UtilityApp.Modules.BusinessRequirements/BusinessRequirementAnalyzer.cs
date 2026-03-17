using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace UtilityApp.Modules.BusinessRequirements
{
    internal sealed class BusinessRequirementAnalysisOptions
    {
        public string WorkbookPath { get; set; }

        public int Year { get; set; }

        public string PeriodType { get; set; }

        public int? PeriodValue { get; set; }

        public bool IncludeProgressDetails { get; set; }
    }

    internal sealed class BusinessRequirementAnalysisResult
    {
        public BusinessRequirementAnalysisResult()
        {
            CategorySummaries = new List<DepartmentCategorySummary>();
            StatusSummaries = new List<DepartmentStatusSummary>();
            OperatorSummaries = new List<OperatorSummary>();
        }

        public string WorkbookPath { get; set; }

        public string PeriodLabel { get; set; }

        public int TotalRecords { get; set; }

        public int FilteredRecords { get; set; }

        public bool ShowOperatorSummary { get; set; }

        public bool ShowProgressDetails { get; set; }

        public IList<DepartmentCategorySummary> CategorySummaries { get; set; }

        public IList<DepartmentStatusSummary> StatusSummaries { get; set; }

        public IList<OperatorSummary> OperatorSummaries { get; set; }

        public string ProgressDetailsText { get; set; }

        public string FormattedReportText { get; set; }
    }

    internal sealed class DepartmentCategorySummary
    {
        public string Department { get; set; }

        public int Category1 { get; set; }

        public int Category2 { get; set; }

        public int Category3 { get; set; }
    }

    internal sealed class DepartmentStatusSummary
    {
        public string Department { get; set; }

        public int Done { get; set; }

        public int UserUat { get; set; }

        public int UatVerify { get; set; }

        public int UatApproval { get; set; }

        public int InProgress { get; set; }
    }

    internal sealed class OperatorSummary
    {
        public string CurrentOperator { get; set; }

        public int InProgress { get; set; }
    }

    internal sealed class BusinessRequirementAnalyzer
    {
        private const string DateColumnName = "Creation Date";
        private const string DepartmentColumnName = "Department";
        private const string CurrentFormColumnName = "Current Form";
        private const string CurrentStageColumnName = "Current Stage";
        private const string BrTypeColumnName = "BR Type";
        private const string CurrentOperatorColumnName = "Current Operator";
        private const string BrTitleColumnName = "BR Title";
        private const string BlankValue = "(blank)";

        private const string RegisterForm = "Business Requirement Register (SCR)";
        private const string HandlingForm = "Business Requirement Handling (SCR)";

        private const string RegisterBusinessOfficerStage = "Business Officer Register";
        private const string RegisterDepartmentVerifyStage = "Department Verify";
        private const string RegisterDepartmentHeadApprovalStage = "Department Head Approval";
        private const string RegisterItdDepartmentHeadApprovalStage = "ITD Department Head Approval";
        private const string RegisterSpecialistStage = "ITD Requirement Specialist";
        private const string RegisterEndStage = "End";
        private const string RegisterTerminatedStage = "Terminated";

        private const string HandlingInProgressStage = "IT Department Officer Handling";
        private const string HandlingDepartmentHeadApprovalStage = "IT Department Head Approval";
        private const string HandlingApplicantTestStage = "Applicant Test";
        private const string HandlingDepartmentVerifyStage = "Department Verify";
        private const string HandlingUserDepartmentHeadApprovalStage = "Department Head Approval";
        private const string HandlingHeadConfirmationStage = "IT Department Head Confirmation";
        private const string HandlingStandbyInstallationStage = "IT Department Offcer standby Installation";
        private const string HandlingInstallationStage = "IT Department Officer Installation";
        private const string HandlingEndStage = "End";
        private const string HandlingTerminatedStage = "Terminated";

        private const string Category1 = "Category 1";
        private const string Category2 = "Category 2";
        private const string Category3 = "Category 3";

        private static readonly string RegisterFormNormalized = NormalizeText(RegisterForm);
        private static readonly string HandlingFormNormalized = NormalizeText(HandlingForm);
        private static readonly string RegisterSpecialistStageNormalized = NormalizeText(RegisterSpecialistStage);
        private static readonly string HandlingInProgressStageNormalized = NormalizeText(HandlingInProgressStage);
        private static readonly string HandlingApplicantTestStageNormalized = NormalizeText(HandlingApplicantTestStage);
        private static readonly string HandlingDepartmentVerifyStageNormalized = NormalizeText(HandlingDepartmentVerifyStage);
        private static readonly string HandlingUserDepartmentHeadApprovalStageNormalized = NormalizeText(HandlingUserDepartmentHeadApprovalStage);
        private static readonly string TerminatedStageNormalized = NormalizeText(RegisterTerminatedStage);
        private static readonly string Category1Normalized = NormalizeText(Category1);
        private static readonly string Category2Normalized = NormalizeText(Category2);
        private static readonly string Category3Normalized = NormalizeText(Category3);

        private static readonly HashSet<string> ValidRegisterStages = new HashSet<string>(StringComparer.Ordinal)
        {
            NormalizeText(RegisterBusinessOfficerStage),
            NormalizeText(RegisterDepartmentVerifyStage),
            NormalizeText(RegisterDepartmentHeadApprovalStage),
            NormalizeText(RegisterItdDepartmentHeadApprovalStage),
            NormalizeText(RegisterSpecialistStage),
            NormalizeText(RegisterEndStage),
            NormalizeText(RegisterTerminatedStage)
        };

        private static readonly HashSet<string> ValidHandlingStages = new HashSet<string>(StringComparer.Ordinal)
        {
            NormalizeText(HandlingInProgressStage),
            NormalizeText(HandlingDepartmentHeadApprovalStage),
            NormalizeText(HandlingApplicantTestStage),
            NormalizeText(HandlingDepartmentVerifyStage),
            NormalizeText(HandlingUserDepartmentHeadApprovalStage),
            NormalizeText(HandlingHeadConfirmationStage),
            NormalizeText(HandlingStandbyInstallationStage),
            NormalizeText(HandlingInstallationStage),
            NormalizeText(HandlingEndStage),
            NormalizeText(HandlingTerminatedStage)
        };

        private static readonly HashSet<string> ValidBrTypes = new HashSet<string>(StringComparer.Ordinal)
        {
            Category1Normalized,
            Category2Normalized,
            Category3Normalized
        };

        private static readonly HashSet<string> RegisterStagesBeforeSpecialist = new HashSet<string>(StringComparer.Ordinal)
        {
            NormalizeText(RegisterBusinessOfficerStage),
            NormalizeText(RegisterDepartmentVerifyStage),
            NormalizeText(RegisterDepartmentHeadApprovalStage),
            NormalizeText(RegisterItdDepartmentHeadApprovalStage)
        };

        private static readonly HashSet<string> DoneHandlingStages = new HashSet<string>(StringComparer.Ordinal)
        {
            NormalizeText(HandlingEndStage),
            NormalizeText(HandlingInstallationStage),
            NormalizeText(HandlingStandbyInstallationStage),
            NormalizeText(HandlingHeadConfirmationStage)
        };

        public BusinessRequirementAnalysisResult Analyze(BusinessRequirementAnalysisOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException("options");
            }

            if (string.IsNullOrWhiteSpace(options.WorkbookPath))
            {
                throw new InvalidOperationException("Workbook path is required.");
            }

            if (string.IsNullOrWhiteSpace(options.PeriodType))
            {
                throw new InvalidOperationException("Period type is required.");
            }

            var workbook = OpenXmlWorkbookReader.ReadFirstWorksheet(options.WorkbookPath);
            EnsureRequiredColumns(workbook, options);

            DateTime startDate;
            DateTime endDate;
            BuildRange(options.Year, options.PeriodType, options.PeriodValue, out startDate, out endDate);

            var filteredRows = new List<WorkbookRow>();
            foreach (var row in workbook.Rows)
            {
                DateTime creationDate;
                if (!TryParseCreationDate(row[DateColumnName], out creationDate))
                {
                    continue;
                }

                if (creationDate >= startDate && creationDate <= endDate)
                {
                    filteredRows.Add(row);
                }
            }

            var departmentOrder = new List<string>();
            var rowsByDepartment = new Dictionary<string, List<WorkbookRow>>(StringComparer.Ordinal);

            foreach (var row in filteredRows)
            {
                var department = NormalizeDisplayValue(row[DepartmentColumnName]);
                List<WorkbookRow> departmentRows;
                if (!rowsByDepartment.TryGetValue(department, out departmentRows))
                {
                    departmentRows = new List<WorkbookRow>();
                    rowsByDepartment[department] = departmentRows;
                    departmentOrder.Add(department);
                }

                departmentRows.Add(row);
            }

            var categorySummaries = new List<DepartmentCategorySummary>();
            var statusSummaries = new List<DepartmentStatusSummary>();
            var inProgressByOperator = new Dictionary<string, int>(StringComparer.Ordinal);
            var progressTitlesByOperator = new Dictionary<string, List<string>>(StringComparer.Ordinal);

            foreach (var department in departmentOrder)
            {
                var departmentRows = rowsByDepartment[department];
                var categorySummary = new DepartmentCategorySummary
                {
                    Department = department
                };
                var statusSummary = new DepartmentStatusSummary
                {
                    Department = department
                };

                foreach (var row in departmentRows)
                {
                    var currentForm = NormalizeText(row[CurrentFormColumnName]);
                    var currentStage = NormalizeText(row[CurrentStageColumnName]);
                    var brType = NormalizeText(row[BrTypeColumnName]);

                    var isRegister = string.Equals(currentForm, RegisterFormNormalized, StringComparison.Ordinal);
                    var isHandling = string.Equals(currentForm, HandlingFormNormalized, StringComparison.Ordinal);
                    var validForm = isRegister || isHandling;
                    var validStage =
                        (isRegister && ValidRegisterStages.Contains(currentStage))
                        || (isHandling && ValidHandlingStages.Contains(currentStage));
                    var validBrType = ValidBrTypes.Contains(brType);

                    var keep = validForm
                        && validStage
                        && validBrType
                        && !string.Equals(currentStage, TerminatedStageNormalized, StringComparison.Ordinal)
                        && !(isRegister && RegisterStagesBeforeSpecialist.Contains(currentStage));

                    if (!keep)
                    {
                        continue;
                    }

                    if (string.Equals(brType, Category1Normalized, StringComparison.Ordinal))
                    {
                        categorySummary.Category1++;
                    }
                    else if (string.Equals(brType, Category2Normalized, StringComparison.Ordinal))
                    {
                        categorySummary.Category2++;
                    }
                    else if (string.Equals(brType, Category3Normalized, StringComparison.Ordinal))
                    {
                        categorySummary.Category3++;
                    }

                    if (isHandling && DoneHandlingStages.Contains(currentStage))
                    {
                        statusSummary.Done++;
                    }

                    if (isHandling && string.Equals(currentStage, HandlingApplicantTestStageNormalized, StringComparison.Ordinal))
                    {
                        statusSummary.UserUat++;
                    }

                    if (isHandling && string.Equals(currentStage, HandlingDepartmentVerifyStageNormalized, StringComparison.Ordinal))
                    {
                        statusSummary.UatVerify++;
                    }

                    if (isHandling && string.Equals(currentStage, HandlingUserDepartmentHeadApprovalStageNormalized, StringComparison.Ordinal))
                    {
                        statusSummary.UatApproval++;
                    }

                    var isInProgress =
                        (isRegister && string.Equals(currentStage, RegisterSpecialistStageNormalized, StringComparison.Ordinal))
                        || (isHandling && string.Equals(currentStage, HandlingInProgressStageNormalized, StringComparison.Ordinal));

                    if (!isInProgress)
                    {
                        continue;
                    }

                    statusSummary.InProgress++;

                    if (string.Equals(options.PeriodType, "m", StringComparison.OrdinalIgnoreCase) || options.IncludeProgressDetails)
                    {
                        var currentOperator = NormalizeDisplayValue(row[CurrentOperatorColumnName]);

                        if (string.Equals(options.PeriodType, "m", StringComparison.OrdinalIgnoreCase))
                        {
                            int count;
                            inProgressByOperator.TryGetValue(currentOperator, out count);
                            inProgressByOperator[currentOperator] = count + 1;
                        }

                        if (options.IncludeProgressDetails)
                        {
                            List<string> titles;
                            if (!progressTitlesByOperator.TryGetValue(currentOperator, out titles))
                            {
                                titles = new List<string>();
                                progressTitlesByOperator[currentOperator] = titles;
                            }

                            titles.Add(NormalizeDisplayValue(row[BrTitleColumnName]));
                        }
                    }
                }

                categorySummaries.Add(categorySummary);
                statusSummaries.Add(statusSummary);
            }

            categorySummaries.Sort(delegate(DepartmentCategorySummary left, DepartmentCategorySummary right)
            {
                return StringComparer.OrdinalIgnoreCase.Compare(left.Department ?? string.Empty, right.Department ?? string.Empty);
            });

            statusSummaries.Sort(delegate(DepartmentStatusSummary left, DepartmentStatusSummary right)
            {
                return StringComparer.OrdinalIgnoreCase.Compare(left.Department ?? string.Empty, right.Department ?? string.Empty);
            });

            var operatorSummaries = inProgressByOperator
                .OrderByDescending(item => item.Value)
                .ThenBy(item => item.Key, StringComparer.Ordinal)
                .Select(item => new OperatorSummary
                {
                    CurrentOperator = item.Key,
                    InProgress = item.Value
                })
                .ToList();

            var progressDetailsText = options.IncludeProgressDetails
                ? BuildProgressDetailsText(progressTitlesByOperator)
                : string.Empty;
            var showOperatorSummary = string.Equals(options.PeriodType, "m", StringComparison.OrdinalIgnoreCase);

            return new BusinessRequirementAnalysisResult
            {
                WorkbookPath = options.WorkbookPath,
                PeriodLabel = BuildPeriodLabel(options.Year, options.PeriodType, options.PeriodValue),
                TotalRecords = workbook.Rows.Count,
                FilteredRecords = filteredRows.Count,
                ShowOperatorSummary = showOperatorSummary,
                ShowProgressDetails = options.IncludeProgressDetails,
                ProgressDetailsText = progressDetailsText,
                FormattedReportText = BuildReportText(
                    workbook.Rows.Count,
                    filteredRows.Count,
                    categorySummaries,
                    statusSummaries,
                    operatorSummaries,
                    progressDetailsText,
                    showOperatorSummary,
                    options.IncludeProgressDetails),
                CategorySummaries = categorySummaries,
                StatusSummaries = statusSummaries,
                OperatorSummaries = operatorSummaries
            };
        }

        private static void EnsureRequiredColumns(WorkbookTable workbook, BusinessRequirementAnalysisOptions options)
        {
            if (!workbook.HasHeader(DateColumnName))
            {
                throw new InvalidOperationException("Creation Date column not found.");
            }

            if (!workbook.HasHeader(DepartmentColumnName))
            {
                throw new InvalidOperationException("Department column not found.");
            }

            if (!workbook.HasHeader(CurrentFormColumnName))
            {
                throw new InvalidOperationException("Current Form column not found.");
            }

            if (!workbook.HasHeader(CurrentStageColumnName))
            {
                throw new InvalidOperationException("Current Stage column not found.");
            }

            if (!workbook.HasHeader(BrTypeColumnName))
            {
                throw new InvalidOperationException("BR Type column not found.");
            }

            if ((string.Equals(options.PeriodType, "m", StringComparison.OrdinalIgnoreCase) || options.IncludeProgressDetails)
                && !workbook.HasHeader(CurrentOperatorColumnName))
            {
                throw new InvalidOperationException("Current Operator column not found.");
            }

            if (options.IncludeProgressDetails && !workbook.HasHeader(BrTitleColumnName))
            {
                throw new InvalidOperationException("BR Title column not found.");
            }
        }

        private static bool TryParseCreationDate(string rawValue, out DateTime creationDate)
        {
            creationDate = DateTime.MinValue;

            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return false;
            }

            var normalized = rawValue.Trim();
            if (normalized.EndsWith(".0", StringComparison.Ordinal))
            {
                normalized = normalized.Substring(0, normalized.Length - 2);
            }

            return DateTime.TryParseExact(
                normalized,
                "yyyyMMdd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out creationDate);
        }

        private static void BuildRange(int year, string periodType, int? periodValue, out DateTime startDate, out DateTime endDate)
        {
            if (string.Equals(periodType, "a", StringComparison.OrdinalIgnoreCase))
            {
                startDate = new DateTime(year, 1, 1);
                endDate = new DateTime(year, 12, 31);
                return;
            }

            if (!periodValue.HasValue)
            {
                throw new InvalidOperationException("Period value is required.");
            }

            if (string.Equals(periodType, "m", StringComparison.OrdinalIgnoreCase))
            {
                startDate = new DateTime(year, periodValue.Value, 1);
                endDate = new DateTime(year, periodValue.Value, DateTime.DaysInMonth(year, periodValue.Value));
                return;
            }

            if (string.Equals(periodType, "q", StringComparison.OrdinalIgnoreCase))
            {
                var startMonth = (periodValue.Value - 1) * 3 + 1;
                var endMonth = startMonth + 2;
                startDate = new DateTime(year, startMonth, 1);
                endDate = new DateTime(year, endMonth, DateTime.DaysInMonth(year, endMonth));
                return;
            }

            var semiAnnualStartMonth = periodValue.Value == 1 ? 1 : 7;
            var semiAnnualEndMonth = periodValue.Value == 1 ? 6 : 12;
            startDate = new DateTime(year, semiAnnualStartMonth, 1);
            endDate = new DateTime(year, semiAnnualEndMonth, DateTime.DaysInMonth(year, semiAnnualEndMonth));
        }

        private static string BuildPeriodLabel(int year, string periodType, int? periodValue)
        {
            if (string.Equals(periodType, "a", StringComparison.OrdinalIgnoreCase))
            {
                return year.ToString(CultureInfo.InvariantCulture);
            }

            if (string.Equals(periodType, "m", StringComparison.OrdinalIgnoreCase))
            {
                return string.Format(CultureInfo.InvariantCulture, "{0} / Month {1}", year, periodValue.GetValueOrDefault());
            }

            if (string.Equals(periodType, "q", StringComparison.OrdinalIgnoreCase))
            {
                return string.Format(CultureInfo.InvariantCulture, "{0} / Quarter {1}", year, periodValue.GetValueOrDefault());
            }

            return string.Format(CultureInfo.InvariantCulture, "{0} / Half {1}", year, periodValue.GetValueOrDefault());
        }

        private static string BuildProgressDetailsText(IDictionary<string, List<string>> progressTitlesByOperator)
        {
            if (progressTitlesByOperator == null || progressTitlesByOperator.Count == 0)
            {
                return "(no in-progress records)";
            }

            var lines = new List<string>();
            foreach (var item in progressTitlesByOperator
                .OrderByDescending(entry => entry.Value == null ? 0 : entry.Value.Count)
                .ThenBy(entry => entry.Key, StringComparer.Ordinal))
            {
                lines.Add(item.Key + ":");

                var titles = item.Value ?? new List<string>();
                for (var index = 0; index < titles.Count; index++)
                {
                    lines.Add(string.Format(CultureInfo.InvariantCulture, "  {0}. {1}", index + 1, titles[index]));
                }
            }

            return string.Join(Environment.NewLine, lines);
        }

        private static string BuildReportText(
            int totalRecords,
            int filteredRecords,
            IList<DepartmentCategorySummary> categorySummaries,
            IList<DepartmentStatusSummary> statusSummaries,
            IList<OperatorSummary> operatorSummaries,
            string progressDetailsText,
            bool includeOperatorSummary,
            bool includeProgressDetails)
        {
            var lines = new List<string>
            {
                string.Format(CultureInfo.InvariantCulture, "Total {0} records", totalRecords),
                string.Format(CultureInfo.InvariantCulture, "After filter {0} records", filteredRecords),
                FormatMatrix(
                    new[] { "Department", "Category 1", "Category 2", "Category 3" },
                    categorySummaries.Select(item => new object[] { item.Department, item.Category1, item.Category2, item.Category3 }).ToList(),
                    new HashSet<int> { 0 }),
                "---------",
                FormatMatrix(
                    new[] { "Department", "Done", "User UAT", "UAT Verify", "UAT Approval", "In progress" },
                    statusSummaries.Select(item => new object[] { item.Department, item.Done, item.UserUat, item.UatVerify, item.UatApproval, item.InProgress }).ToList(),
                    new HashSet<int> { 0 })
            };

            if (includeOperatorSummary)
            {
                lines.Add("------------");
                lines.Add(FormatMatrix(
                    new[] { "Current Operator", "In progress" },
                    operatorSummaries.Select(item => new object[] { item.CurrentOperator, item.InProgress }).ToList(),
                    new HashSet<int> { 0 }));
            }

            if (includeProgressDetails)
            {
                lines.Add("------------");
                lines.Add("In Progress BR Title by Current Operator");
                lines.Add(string.IsNullOrWhiteSpace(progressDetailsText) ? "(no in-progress records)" : progressDetailsText);
            }

            return string.Join(Environment.NewLine, lines);
        }

        private static string FormatMatrix(IList<string> headers, IList<object[]> rows, ISet<int> leftAlignColumns)
        {
            var widths = new int[headers.Count];
            for (var index = 0; index < headers.Count; index++)
            {
                widths[index] = headers[index] == null ? 0 : headers[index].Length;
            }

            var textRows = new List<string[]>();
            foreach (var row in rows)
            {
                var textRow = new string[headers.Count];
                for (var index = 0; index < headers.Count; index++)
                {
                    var text = row != null && index < row.Length && row[index] != null
                        ? Convert.ToString(row[index], CultureInfo.InvariantCulture)
                        : string.Empty;

                    textRow[index] = text ?? string.Empty;
                    if (textRow[index].Length > widths[index])
                    {
                        widths[index] = textRow[index].Length;
                    }
                }

                textRows.Add(textRow);
            }

            var separatorBuilder = new StringBuilder();
            separatorBuilder.Append("+");
            for (var index = 0; index < widths.Length; index++)
            {
                separatorBuilder.Append(new string('-', widths[index] + 2));
                separatorBuilder.Append("+");
            }

            var separator = separatorBuilder.ToString();
            var lines = new List<string>
            {
                separator,
                FormatMatrixRow(headers.ToArray(), widths, leftAlignColumns),
                separator
            };

            foreach (var textRow in textRows)
            {
                lines.Add(FormatMatrixRow(textRow, widths, leftAlignColumns));
            }

            lines.Add(separator);
            return string.Join(Environment.NewLine, lines);
        }

        private static string FormatMatrixRow(IList<string> row, IList<int> widths, ISet<int> leftAlignColumns)
        {
            var builder = new StringBuilder();
            builder.Append("|");

            for (var index = 0; index < widths.Count; index++)
            {
                var cell = row[index] ?? string.Empty;
                builder.Append(" ");
                builder.Append(leftAlignColumns.Contains(index)
                    ? cell.PadRight(widths[index])
                    : cell.PadLeft(widths[index]));
                builder.Append(" |");
            }

            return builder.ToString();
        }

        private static string NormalizeDisplayValue(string value)
        {
            var normalized = (value ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(normalized) ? BlankValue : normalized;
        }

        private static string NormalizeText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var parts = value
                .Trim()
                .Split((char[])null, StringSplitOptions.RemoveEmptyEntries);

            return string.Join(" ", parts).ToLowerInvariant();
        }
    }
}
