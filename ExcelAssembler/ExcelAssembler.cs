using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml.XPath;
using OfficeOpenXml.Style;

namespace ExcelAssembler;

// Core model to represent a repeat block in the sheet

public class ExcelAssembler(ExcelAssemblerOptions assemblerOptions)
{
    private static readonly Regex ContentRegex = new(@"<Content\s+Select\s*=\s*""([^""]+)""\s*/>", RegexOptions.IgnoreCase);
    private static readonly Regex RepeatRegex = new(@"<Repeat\s+Select\s*=\s*""([^""]+)""\s*/>", RegexOptions.IgnoreCase);
    private static readonly Regex EndRepeatRegex = new(@"<EndRepeat\s*/>", RegexOptions.IgnoreCase);
    private static readonly Regex NumericRegex = new(@"[^\d\.\-]", RegexOptions.IgnoreCase);

    internal record RepeatBlock(int StartRow, int TemplateRow, int FormatRow, int EndRow, string XPath);

    public Stream ProcessTemplate(Stream templateStream, string xmlData)
    {
        var xml = XElement.Parse(xmlData);

        using var package = new ExcelPackage(templateStream);

        foreach (var worksheet in package.Workbook.Worksheets)
        {
            ProcessWorksheet(xml, worksheet);
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

    private void ProcessWorksheet(XElement xml, ExcelWorksheet worksheet)
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
            if (TryStartRepeat(worksheet, dim, row, out var repeatBlock))
            {
                var blockShift = ProcessRepeat(worksheet, dim, xml, repeatBlock);

                rowsToDelete.Add(repeatBlock.StartRow);
                rowsToDelete.Add(repeatBlock.TemplateRow);
                rowsToDelete.Add(repeatBlock.FormatRow);
                rowsToDelete.Add(repeatBlock.EndRow + blockShift);

                Console.WriteLine($"Repeat block found from row {repeatBlock.StartRow} to {repeatBlock.EndRow}");

                // Skip ahead to the end of the repeat block before continuing
                row = repeatBlock.EndRow;
                continue;
            }

            if (TryProcessContentRow(worksheet, dim, row, xml))
            {
                // If there were content tags in this row, we just filled in the
                // row below it and now it can be deleted
                rowsToDelete.Add(row);
            }
        }

        // Delete rows bottom-up to avoid index shifting
        foreach (var row in rowsToDelete.OrderByDescending(r => r))
        {
            worksheet.DeleteRow(row);
        }
    }

    private bool TryProcessContentRow(ExcelWorksheet worksheet, ExcelAddressBase dim, int row,
        XElement xml)
    {
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
                PopulateCell(targetCell, result, fixedPath);
            }
        }

        return rowContainsContentTags;
    }

    private static bool TryStartRepeat(ExcelWorksheet worksheet, ExcelAddressBase dim, int startRepeatRow,
        out RepeatBlock repeatBlock)
    {
        if (!TryFindStartRepeatInRow(worksheet, dim, startRepeatRow, out var repeatXPath))
        {
            repeatBlock = null!;
            return false;
        }

        var templateRow = startRepeatRow + 1;
        var formatRow = templateRow + 1;

        // Now locate the <EndRepeat />
        int? endRepeatRow = null;
        for (var scanRow = formatRow + 1; scanRow <= dim.End.Row; scanRow++)
        {
            for (var scanCol = dim.Start.Column; scanCol <= dim.End.Column; scanCol++)
            {
                var scanCell = worksheet.Cells[scanRow, scanCol];
                if (EndRepeatRegex.IsMatch(scanCell.Text))
                {
                    endRepeatRow = scanRow;

                    var blockLength = endRepeatRow.Value - startRepeatRow - 1;

                    if (blockLength != 2)
                    {
                        throw new($"Repeat block at row {startRepeatRow} must contain exactly one template row and one format row. Found {blockLength} rows.");
                    }

                    break;
                }

                if (RepeatRegex.IsMatch(scanCell.Text))
                {
                    throw new($"Nested <Repeat> detected at row {scanRow}. Nested repeats are not supported.");
                }
            }

            if (endRepeatRow != null)
            {
                break;
            }
        }

        if (endRepeatRow == null)
        {
            throw new($"<EndRepeat /> not found after row {startRepeatRow}");
        }

        repeatBlock = new(startRepeatRow, templateRow, formatRow, endRepeatRow.Value, repeatXPath);
        return true;
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

    private int ProcessRepeat(ExcelWorksheet worksheet, ExcelAddressBase dim, XElement xml, RepeatBlock repeat)
    {
        var fixedRepeatXPath = FixRelativeXPath(repeat.XPath);

        // Check if the XPath exists in the XML
        var repeatResult = xml.XPathSelectElement(fixedRepeatXPath);
        if (repeatResult == null)
        {
            switch (assemblerOptions.MissingXmlDataBehaviour)
            {
                case MissingXmlDataBehaviour.ThrowException:
                    throw new($"Unresolved XPath: {fixedRepeatXPath}");
                case MissingXmlDataBehaviour.ShowPlaceholder:
                    // Stick a helpful error message in the cell
                    var cell = worksheet.Cells[repeat.StartRow, dim.Start.Column];
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                    cell.Value = $"[UNRESOLVED: {fixedRepeatXPath}]";
                    return 0;
            }
        }

        var itemNodes = xml.XPathSelectElements(fixedRepeatXPath).ToList();

        var blockShift = 0;

        foreach (var itemNode in itemNodes)
        {
            // Clone the formatting row
            worksheet.InsertRow(repeat.EndRow + blockShift, 1);
            var targetRow = repeat.EndRow + blockShift;

            // Copy format from formatRow to targetRow
            worksheet.Cells[repeat.FormatRow, dim.Start.Column, repeat.FormatRow, dim.End.Column]
                .Copy(worksheet.Cells[targetRow, dim.Start.Column]);

            // Fill in values based on templateRow
            for (var col = dim.Start.Column; col <= dim.End.Column; col++)
            {
                var templateCell = worksheet.Cells[repeat.TemplateRow, col];
                var cellText = templateCell.Text;

                if (ContentRegex.TryMatchGroups(cellText, out var matchGroups))
                {
                    var relativeXPath = matchGroups[1].Value;
                    var fixedPath = FixRelativeXPath(relativeXPath);

                    var result = itemNode.XPathSelectElement(fixedPath);
                    var targetCell = worksheet.Cells[targetRow, col];
                    PopulateCell(targetCell, result, relativeXPath);
                }
            }

            blockShift++; // we’ve added a row, so shift insert point down
        }

        return blockShift;
    }

    private void PopulateCell(ExcelRange targetCell, XElement? result, string xPath)
    {
        if (result == null)
        {
            switch (assemblerOptions.MissingXmlDataBehaviour)
            {
                case MissingXmlDataBehaviour.ShowPlaceholder:
                    targetCell.Value = $"[UNRESOLVED: {xPath}]";
                    targetCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    targetCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                    return;
                case MissingXmlDataBehaviour.ThrowException:
                default:
                    throw new($"Unresolved XPath: {xPath}");
            }
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