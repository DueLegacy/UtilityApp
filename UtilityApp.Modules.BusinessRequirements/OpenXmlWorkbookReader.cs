using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace UtilityApp.Modules.BusinessRequirements
{
    internal sealed class WorkbookTable
    {
        private readonly HashSet<string> _headers;

        public WorkbookTable(IEnumerable<string> headers, IEnumerable<WorkbookRow> rows)
        {
            Headers = new List<string>(headers ?? Enumerable.Empty<string>());
            Rows = new List<WorkbookRow>(rows ?? Enumerable.Empty<WorkbookRow>());
            _headers = new HashSet<string>(Headers.Where(header => !string.IsNullOrWhiteSpace(header)), StringComparer.OrdinalIgnoreCase);
        }

        public IList<string> Headers { get; private set; }

        public IList<WorkbookRow> Rows { get; private set; }

        public bool HasHeader(string header)
        {
            return !string.IsNullOrWhiteSpace(header) && _headers.Contains(header);
        }
    }

    internal sealed class WorkbookRow
    {
        private readonly Dictionary<string, string> _values;

        public WorkbookRow(IDictionary<string, string> values)
        {
            _values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (values == null)
            {
                return;
            }

            foreach (var item in values)
            {
                if (!string.IsNullOrWhiteSpace(item.Key))
                {
                    _values[item.Key] = item.Value ?? string.Empty;
                }
            }
        }

        public string this[string columnName]
        {
            get
            {
                string value;
                return _values.TryGetValue(columnName ?? string.Empty, out value) ? value : string.Empty;
            }
        }
    }

    internal static class OpenXmlWorkbookReader
    {
        private static readonly XNamespace SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace OfficeDocumentRelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public static WorkbookTable ReadFirstWorksheet(string workbookPath)
        {
            if (string.IsNullOrWhiteSpace(workbookPath))
            {
                throw new InvalidOperationException("Workbook path is required.");
            }

            if (!File.Exists(workbookPath))
            {
                throw new InvalidOperationException("Export.xlsx not found.");
            }

            using (var package = Package.Open(workbookPath, FileMode.Open, FileAccess.Read))
            {
                var workbookUri = new Uri("/xl/workbook.xml", UriKind.Relative);
                if (!package.PartExists(workbookUri))
                {
                    throw new InvalidOperationException("Workbook content is invalid.");
                }

                var sharedStrings = LoadSharedStrings(package);
                var workbookDocument = LoadXml(package.GetPart(workbookUri));
                var workbookRelationships = LoadWorkbookRelationships(package, workbookUri);

                var firstSheet = workbookDocument
                    .Descendants(SpreadsheetNamespace + "sheet")
                    .FirstOrDefault();

                if (firstSheet == null)
                {
                    throw new InvalidOperationException("Workbook does not contain any worksheets.");
                }

                var relationshipId = (string)firstSheet.Attribute(OfficeDocumentRelationshipsNamespace + "id");
                if (string.IsNullOrWhiteSpace(relationshipId) || !workbookRelationships.ContainsKey(relationshipId))
                {
                    throw new InvalidOperationException("Workbook worksheet relationship is invalid.");
                }

                var worksheetUri = workbookRelationships[relationshipId];
                if (!package.PartExists(worksheetUri))
                {
                    throw new InvalidOperationException("Worksheet part is missing.");
                }

                var worksheetDocument = LoadXml(package.GetPart(worksheetUri));
                return ReadWorksheetRows(worksheetDocument, sharedStrings);
            }
        }

        private static Dictionary<string, Uri> LoadWorkbookRelationships(Package package, Uri workbookUri)
        {
            var relationships = new Dictionary<string, Uri>(StringComparer.Ordinal);
            var relationshipsUri = PackUriHelper.GetRelationshipPartUri(workbookUri);

            if (!package.PartExists(relationshipsUri))
            {
                return relationships;
            }

            var relationshipsDocument = LoadXml(package.GetPart(relationshipsUri));
            foreach (var relationship in relationshipsDocument.Descendants(RelationshipsNamespace + "Relationship"))
            {
                var id = (string)relationship.Attribute("Id");
                var target = (string)relationship.Attribute("Target");
                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(target))
                {
                    continue;
                }

                relationships[id] = PackUriHelper.ResolvePartUri(workbookUri, new Uri(target, UriKind.Relative));
            }

            return relationships;
        }

        private static IList<string> LoadSharedStrings(Package package)
        {
            var sharedStringsUri = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
            if (!package.PartExists(sharedStringsUri))
            {
                return new List<string>();
            }

            var sharedStringsDocument = LoadXml(package.GetPart(sharedStringsUri));
            return sharedStringsDocument
                .Descendants(SpreadsheetNamespace + "si")
                .Select(ReadSharedStringValue)
                .ToList();
        }

        private static WorkbookTable ReadWorksheetRows(XDocument worksheetDocument, IList<string> sharedStrings)
        {
            var rowMaps = new List<Dictionary<int, string>>();
            var sheetData = worksheetDocument.Descendants(SpreadsheetNamespace + "sheetData").FirstOrDefault();
            if (sheetData != null)
            {
                foreach (var rowElement in sheetData.Elements(SpreadsheetNamespace + "row"))
                {
                    rowMaps.Add(ReadRowMap(rowElement, sharedStrings));
                }
            }

            if (rowMaps.Count == 0)
            {
                return new WorkbookTable(new string[0], new WorkbookRow[0]);
            }

            var headerMap = rowMaps[0];
            var maxHeaderIndex = headerMap.Count == 0 ? -1 : headerMap.Keys.Max();
            var headers = new List<string>();

            for (var columnIndex = 0; columnIndex <= maxHeaderIndex; columnIndex++)
            {
                string headerValue;
                headerMap.TryGetValue(columnIndex, out headerValue);
                headers.Add((headerValue ?? string.Empty).Trim());
            }

            var rows = new List<WorkbookRow>();
            for (var rowIndex = 1; rowIndex < rowMaps.Count; rowIndex++)
            {
                var rowMap = rowMaps[rowIndex];
                var rowValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                for (var columnIndex = 0; columnIndex < headers.Count; columnIndex++)
                {
                    var header = headers[columnIndex];
                    if (string.IsNullOrWhiteSpace(header))
                    {
                        continue;
                    }

                    string cellValue;
                    rowMap.TryGetValue(columnIndex, out cellValue);
                    rowValues[header] = cellValue ?? string.Empty;
                }

                rows.Add(new WorkbookRow(rowValues));
            }

            return new WorkbookTable(headers, rows);
        }

        private static Dictionary<int, string> ReadRowMap(XElement rowElement, IList<string> sharedStrings)
        {
            var rowMap = new Dictionary<int, string>();
            var fallbackColumnIndex = 0;

            foreach (var cellElement in rowElement.Elements(SpreadsheetNamespace + "c"))
            {
                var cellReference = (string)cellElement.Attribute("r");
                var columnIndex = GetColumnIndex(cellReference);
                if (columnIndex < 0)
                {
                    columnIndex = fallbackColumnIndex;
                }

                rowMap[columnIndex] = ReadCellValue(cellElement, sharedStrings);
                fallbackColumnIndex = columnIndex + 1;
            }

            return rowMap;
        }

        private static string ReadCellValue(XElement cellElement, IList<string> sharedStrings)
        {
            var cellType = (string)cellElement.Attribute("t");
            if (string.Equals(cellType, "inlineStr", StringComparison.OrdinalIgnoreCase))
            {
                return ReadInlineStringValue(cellElement);
            }

            var valueElement = cellElement.Element(SpreadsheetNamespace + "v");
            var rawValue = valueElement == null ? string.Empty : valueElement.Value;

            if (string.Equals(cellType, "s", StringComparison.OrdinalIgnoreCase))
            {
                int index;
                if (int.TryParse(rawValue, out index) && index >= 0 && index < sharedStrings.Count)
                {
                    return sharedStrings[index];
                }

                return string.Empty;
            }

            return rawValue ?? string.Empty;
        }

        private static string ReadInlineStringValue(XElement cellElement)
        {
            var inlineString = cellElement.Element(SpreadsheetNamespace + "is");
            if (inlineString == null)
            {
                return string.Empty;
            }

            return string.Concat(inlineString
                .Descendants(SpreadsheetNamespace + "t")
                .Select(item => item.Value));
        }

        private static string ReadSharedStringValue(XElement sharedStringItem)
        {
            return string.Concat(sharedStringItem
                .Descendants(SpreadsheetNamespace + "t")
                .Select(item => item.Value));
        }

        private static int GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrWhiteSpace(cellReference))
            {
                return -1;
            }

            var columnIndex = 0;
            var hasLetters = false;

            for (var index = 0; index < cellReference.Length; index++)
            {
                var character = char.ToUpperInvariant(cellReference[index]);
                if (character < 'A' || character > 'Z')
                {
                    break;
                }

                hasLetters = true;
                columnIndex = (columnIndex * 26) + (character - 'A' + 1);
            }

            return hasLetters ? columnIndex - 1 : -1;
        }

        private static XDocument LoadXml(PackagePart part)
        {
            using (var stream = part.GetStream(FileMode.Open, FileAccess.Read))
            {
                return XDocument.Load(stream);
            }
        }
    }
}
