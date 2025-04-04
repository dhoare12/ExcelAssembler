using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml.XPath;

namespace ExcelAssembler;

public class ExcelAssembler
{
    private static readonly Regex ContentRegex = new Regex(@"<Content\s+Select\s*=\s*""([^""]+)""\s*/>", RegexOptions.IgnoreCase);
    private static readonly Regex RepeatRegex = new Regex(@"<Repeat\s+Select\s*=\s*""([^""]+)""\s*/>", RegexOptions.IgnoreCase);
    private static readonly Regex EndRepeatRegex = new Regex(@"<EndRepeat\s*/>", RegexOptions.IgnoreCase);

    public void ProcessTemplate(string templatePath, string xmlPath, string outputPath)
    {
        var xml = XElement.Load(xmlPath);

        using var package = new ExcelPackage(new FileInfo(templatePath));
        ProcessWithFormatting(package, xml);

        package.SaveAs(new FileInfo(outputPath));
    }

    public void ProcessWithFormatting(ExcelPackage package, XElement xml)
    {
        var rowsToDelete = new List<int>();

        foreach (var worksheet in package.Workbook.Worksheets)
        {
            var dim = worksheet.Dimension;
            if (dim == null) continue;

            for (var row = dim.Start.Row; row <= dim.End.Row; row++)
            {
                var foundTag = false;
                var isRepeatHandled = false;

                for (var col = dim.Start.Column; col <= dim.End.Column; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    var repeatMatch = RepeatRegex.Match(cell.Text);

                    if (repeatMatch.Success)
                    {
                        string repeatXPath = repeatMatch.Groups[1].Value;
                        int repeatRow = row;
                        int templateRow = repeatRow + 1;
                        int formatRow = templateRow + 1;

                        // Now locate the endRepeatRow
                        int endRepeatRow = -1;
                        for (int scanRow = formatRow + 1; scanRow <= dim.End.Row; scanRow++)
                        {
                            for (int scanCol = dim.Start.Column; scanCol <= dim.End.Column; scanCol++)
                            {
                                var scanCell = worksheet.Cells[scanRow, scanCol];
                                if (EndRepeatRegex.IsMatch(scanCell.Text))
                                {
                                    endRepeatRow = scanRow;
                                    break;
                                }
                            }

                            if (endRepeatRow != -1)
                                break;
                        }

                        if (endRepeatRow == -1)
                        {
                            throw new Exception($"<EndRepeat /> not found after row {repeatRow}");
                        }

                        

                        var blockShift = ProcessRepeat(xml, repeatXPath, endRepeatRow, worksheet, formatRow, dim, templateRow);

                        rowsToDelete.Add(repeatRow);
                        rowsToDelete.Add(templateRow);
                        rowsToDelete.Add(formatRow);
                        rowsToDelete.Add(endRepeatRow + blockShift);

                        Console.WriteLine($"Repeat block found from row {repeatRow} to {endRepeatRow}");

                        // Optional: skip ahead past the end of the block
                        row = endRepeatRow;
                        break;
                    }
                }

                if (isRepeatHandled)
                    continue;

                for (var col = dim.Start.Column; col <= dim.End.Column; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    var contentMatch = ContentRegex.Match(cell.Text);

                    if (contentMatch.Success)
                    {
                        foundTag = true;

                        var xpath = contentMatch.Groups[1].Value;
                        var fixedPath = FixRelativeXPath(xpath);

                        var result = xml.XPathSelectElement(fixedPath);
                        if (result != null)
                        {
                            // Write into the row below, same column
                            PopulateCell(worksheet, row, col, result);
                        }
                        else
                        {
                            worksheet.Cells[row + 1, col].Value = $"[UNRESOLVED: {xpath}]";
                        }
                    }
                }

                if (foundTag)
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
    }

    private int ProcessRepeat(XElement xml, string repeatXPath, int endRepeatRow, ExcelWorksheet worksheet, int formatRow,
        ExcelAddressBase dim, int templateRow)
    {
        var fixedRepeatXPath = FixRelativeXPath(repeatXPath);
        var itemNodes = xml.XPathSelectElements(fixedRepeatXPath).ToList();

        int insertAt = endRepeatRow; // We'll insert new rows just before EndRepeat
        int blockShift = 0;

        foreach (var itemNode in itemNodes)
        {
            // Clone the formatting row
            worksheet.InsertRow(insertAt + blockShift, 1);
            var targetRow = insertAt + blockShift;

            // Copy format from formatRow to targetRow
            worksheet.Cells[formatRow, dim.Start.Column, formatRow, dim.End.Column]
                .Copy(worksheet.Cells[targetRow, dim.Start.Column]);

            // Fill in values based on templateRow
            for (int col = dim.Start.Column; col <= dim.End.Column; col++)
            {
                var templateCell = worksheet.Cells[templateRow, col];
                var match = ContentRegex.Match(templateCell.Text);

                if (match.Success)
                {
                    string relativeXPath = match.Groups[1].Value;
                    string fixedPath = FixRelativeXPath(relativeXPath);

                    var result = itemNode.XPathSelectElement(fixedPath);
                    var targetCell = worksheet.Cells[targetRow, col];

                    if (result != null)
                    {
                        var format = targetCell.Style.Numberformat.Format?.ToLowerInvariant() ?? "";
                        if (format.Contains("0") || format.Contains("#"))
                        {
                            string numeric = Regex.Replace(result.Value, @"[^\d\.\-]", "");
                            if (decimal.TryParse(numeric, out var number))
                                targetCell.Value = number;
                            else
                                targetCell.Value = result.Value;
                        }
                        else
                        {
                            targetCell.Value = result.Value;
                        }
                    }
                    else
                    {
                        targetCell.Value = $"[UNRESOLVED: {relativeXPath}]";
                    }
                }
            }

            blockShift++; // we’ve added a row, so shift insert point down
        }

        return blockShift;
    }

    private static void PopulateCell(ExcelWorksheet worksheet, int row, int col, XElement result)
    {
        var targetCell = worksheet.Cells[row + 1, col];
        var format = targetCell.Style.Numberformat.Format?.ToLowerInvariant() ?? "";

        var rawValue = result.Value;

        if (format.Contains("0") || format.Contains("#")) // likely numeric
        {
            var numericString = Regex.Replace(rawValue, @"[^\d\.\-]", ""); // keep digits, dot, minus
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

    private void ProcessWithoutFormatting(ExcelPackage package, XElement xml)
    {
        foreach (var worksheet in package.Workbook.Worksheets)
        {
            var dimension = worksheet.Dimension;
            if (dimension == null) continue;

            for (var row = dimension.Start.Row; row <= dimension.End.Row; row++)
            {
                for (var col = dimension.Start.Column; col <= dimension.End.Column; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    var cellText = cell.Text;

                    var match = ContentRegex.Match(cellText);
                    if (match.Success)
                    {
                        var xpath = match.Groups[1].Value;
                        var fixedPath = FixRelativeXPath(xpath);
                        var result = xml.XPathSelectElement(fixedPath);
                        if (result != null)
                        {
                            cell.Value = result.Value;
                        }
                        else
                        {
                            // Optional: mark as unresolved
                            cell.Value = $"[UNRESOLVED: {xpath}]";
                        }
                    }
                }
            }
        }
    }

    private string FixRelativeXPath(string xpath)
    {
        if (xpath.StartsWith("./"))
            return xpath.Substring(2); // Remove leading ./
        if (xpath.StartsWith("/"))
            return xpath.Substring(1); // Remove leading /
        return xpath;
    }
}