using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml.XPath;
using OfficeOpenXml.ConditionalFormatting;

namespace ExcelAssembler;

public static partial class ExcelAssembler
{
    private static readonly Regex ContentRegex = new Regex(@"<Content\s+Select\s*=\s*""([^""]+)""\s*/>", RegexOptions.IgnoreCase);
    private static readonly Regex RepeatRegex = new Regex(@"<Repeat\s+Select\s*=\s*""([^""]+)""\s*/>", RegexOptions.IgnoreCase);
    private static readonly Regex EndRepeatRegex = new Regex(@"<EndRepeat\s*/>", RegexOptions.IgnoreCase);
    private static readonly Regex NumericRegex = new Regex(@"[^\d\.\-]", RegexOptions.IgnoreCase);

    public static Stream ProcessTemplate(Stream templateStream, string xmlData, bool suppressMissingXml = false)
    {
        var xml = XElement.Parse(xmlData);

        using var package = new ExcelPackage(templateStream);

        foreach (var worksheet in package.Workbook.Worksheets)
        {
            ProcessWorksheet(xml, worksheet, suppressMissingXml);
        }

        var memoryStream = new MemoryStream();
        package.SaveAs(memoryStream);
        memoryStream.Position = 0; // Reset stream position to the beginning
        return memoryStream;
    }

    private static void ValidateTags(ExcelWorksheet worksheet, ExcelAddressBase dim)
    {
        for (var row = dim.Start.Row; row <= dim.End.Row; row++)
        {
            for (var col = dim.Start.Column; col <= dim.End.Column; col++)
            {
                var cell = worksheet.Cells[row, col];
                var cellText = cell.Text;

                if (cellText.StartsWith("<Content") && !ContentRegex.IsMatch(cellText))
                {
                    throw new($"Invalid <Content> tag at row {row}, column {col}. Expected format: <Content Select=\"...\" />");
                }
                if (cellText.StartsWith("<Repeat") && !RepeatRegex.IsMatch(cellText))
                {
                    throw new($"Invalid <Repeat> tag at row {row}, column {col}. Expected format: <Repeat Select=\"...\" />");
                }
            }
        }
    }

    private static void ProcessWorksheet(XElement xml, ExcelWorksheet worksheet, bool suppressMissingXml)
    {
        var rowsToDelete = new List<int>();

        var dim = worksheet.Dimension;
        if (dim == null)
        {
            return;
        }

        ValidateTags(worksheet, dim);

        for (var row = dim.Start.Row; row <= dim.End.Row; row++)
        {
            if (TryFindStartRepeatInRow(worksheet, dim, row, out var repeatXPath))
            {
                var repeatRow = row;
                var templateRow = repeatRow + 1;
                var formatRow = templateRow + 1;

                // Now locate the endRepeatRow
                int? endRepeatRow = null;
                for (var scanRow = formatRow + 1; scanRow <= dim.End.Row; scanRow++)
                {
                    for (var scanCol = dim.Start.Column; scanCol <= dim.End.Column; scanCol++)
                    {
                        var scanCell = worksheet.Cells[scanRow, scanCol];
                        if (EndRepeatRegex.IsMatch(scanCell.Text))
                        {
                            endRepeatRow = scanRow;

                            var blockLength = endRepeatRow.Value - repeatRow - 1;

                            if (blockLength != 2)
                            {
                                throw new($"Repeat block at row {repeatRow} must contain exactly one template row and one format row. Found {blockLength} rows.");
                            }

                            break;
                        }
                    }

                    if (endRepeatRow != null)
                    {
                        break;
                    }
                }

                if (endRepeatRow == null)
                {
                    throw new($"<EndRepeat /> not found after row {repeatRow}");
                }

                var blockShift = ProcessRepeat(xml, repeatXPath, endRepeatRow.Value, worksheet, formatRow, dim, templateRow, suppressMissingXml);

                rowsToDelete.Add(repeatRow);
                rowsToDelete.Add(templateRow);
                rowsToDelete.Add(formatRow);
                rowsToDelete.Add(endRepeatRow.Value + blockShift);

                Console.WriteLine($"Repeat block found from row {repeatRow} to {endRepeatRow}");

                row = endRepeatRow.Value;
                continue;
            }

            var rowContainsContentTags = false;

            for (var col = dim.Start.Column; col <= dim.End.Column; col++)
            {
                var cell = worksheet.Cells[row, col];

                if (ContentRegex.TryMatchGroups(cell.Text, out var contentMatchGroups))
                {
                    rowContainsContentTags = true;

                    var xpath = contentMatchGroups[1].Value;
                    var fixedPath = FixRelativeXPath(xpath);

                    var result = xml.XPathSelectElement(fixedPath);
                    var targetCell = worksheet.Cells[row + 1, col];
                    PopulateCell(targetCell, result, fixedPath, suppressMissingXml);
                }
            }

            if (rowContainsContentTags)
            {
                rowsToDelete.Add(row);
            }
        }

        // Delete rows bottom-up to avoid index shifting
        foreach (var row in rowsToDelete.OrderByDescending(r => r))
        {
            worksheet.DeleteRow(row);
        }
    }

    private static bool TryFindStartRepeatInRow(ExcelWorksheet worksheet, ExcelAddressBase dim, int row, out string xPath)
    {
        for (var col = dim.Start.Column; col <= dim.End.Column; col++)
        {
            var cell = worksheet.Cells[row, col];

            if (RepeatRegex.TryMatchGroups(cell.Text, out var repeatMatchGroups))
            {
                xPath = repeatMatchGroups[1].Value;
                return true;
            }
        }

        xPath = null!;
        return false;
    }

    private static int ProcessRepeat(XElement xml, string repeatXPath, int endRepeatRow, ExcelWorksheet worksheet, int formatRow,
        ExcelAddressBase dim, int templateRow, bool suppressMissingXml)
    {
        var fixedRepeatXPath = FixRelativeXPath(repeatXPath);

        if (!suppressMissingXml)
        {
            // Check if the XPath exists in the XML
            var result = xml.XPathSelectElement(fixedRepeatXPath);
            if (result == null)
            {
                throw new($"Unresolved XPath: {fixedRepeatXPath}");
            }
        }
        var itemNodes = xml.XPathSelectElements(fixedRepeatXPath).ToList();

        var blockShift = 0;

        foreach (var itemNode in itemNodes)
        {
            // Clone the formatting row
            worksheet.InsertRow(endRepeatRow + blockShift, 1);
            var targetRow = endRepeatRow + blockShift;

            // Copy format from formatRow to targetRow
            worksheet.Cells[formatRow, dim.Start.Column, formatRow, dim.End.Column]
                .Copy(worksheet.Cells[targetRow, dim.Start.Column]);

            // Fill in values based on templateRow
            for (var col = dim.Start.Column; col <= dim.End.Column; col++)
            {
                var templateCell = worksheet.Cells[templateRow, col];
                var cellText = templateCell.Text;

                if (RepeatRegex.IsMatch(cellText))
                {
                    throw new($"Nested <Repeat> detected at row {formatRow}. Nested repeats are not supported.");
                }

                if (EndRepeatRegex.IsMatch(cellText))
                {
                    throw new($"Unexpected <EndRepeat /> inside repeat block at row {formatRow}.");
                }

                if (ContentRegex.TryMatchGroups(cellText, out var matchGroups))
                {
                    var relativeXPath = matchGroups[1].Value;
                    var fixedPath = FixRelativeXPath(relativeXPath);

                    var result = itemNode.XPathSelectElement(fixedPath);
                    var targetCell = worksheet.Cells[targetRow, col];
                    PopulateCell(targetCell, result, relativeXPath, suppressMissingXml);
                }
            }

            blockShift++; // we’ve added a row, so shift insert point down
        }

        return blockShift;
    }

    private static void PopulateCell(ExcelRange targetCell, XElement? result, string xPath, bool suppressMissingXml)
    {
        if (result == null)
        {
            if (!suppressMissingXml)
            {
                throw new Exception($"Unresolved XPath: {xPath}");
            }
            targetCell.Value = $"[UNRESOLVED: {xPath}]";
            return;
        }

        var format = targetCell.Style.Numberformat.Format?.ToLowerInvariant() ?? "";

        var rawValue = result.Value;

        if (format.Contains('0') || format.Contains('#')) // likely numeric
        {
            var numericString = NumericRegex.Replace(rawValue, ""); // keep digits, dot, minus
            if (decimal.TryParse(numericString, out var number))
            {
                targetCell.Value = number;
            }
            else
            {
                targetCell.Value = rawValue; // fallback if parse fails
            }
        }
        else
        {
            targetCell.Value = rawValue;
        }
    }

    private static string FixRelativeXPath(string xpath)
    {
        if (xpath.StartsWith("./"))
        {
            return xpath.Substring(2); // Remove leading ./
        }

        if (xpath.StartsWith("/"))
        {
            return xpath.Substring(1); // Remove leading /
        }

        return xpath;
    }
}

public static class RegexExtensions
{
    public static bool TryMatchGroups(this Regex regex, string input, out GroupCollection groups)
    {
        var matches = regex.Match(input);

        groups = matches.Groups;
        return matches.Success;
    }
}