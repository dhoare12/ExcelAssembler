using NUnit.Framework;
using OfficeOpenXml;
using Shouldly;

namespace ExcelAssembler.Tests
{
    public class UnitTest1
    {
        [Test]
        [TestCase("Content_Simple")]
        [TestCase("Content_Repeat")]
        [TestCase("Malformed_Tag")]
        [TestCase("Not_Found_XML")]
        public void TestCasesShouldRun(string testCaseName)
        {
            var templatePath = Path.Combine(testCaseName, $"{testCaseName}.xlsx");
            var xmlPath = "TestInput.xml";
            var expectedPath = Path.Combine(testCaseName, $"{testCaseName}_ExpectedOutput.xlsx");
            var expectedErrorPath = Path.Combine(testCaseName, $"{testCaseName}_ExpectedError.txt");

            Stream outputStream;
            try
            {
                outputStream = ExcelAssembler.ProcessTemplate(File.OpenRead(templatePath), File.ReadAllText(xmlPath));
            }
            catch (Exception ex)
            {
                if (!File.Exists(expectedErrorPath))
                {
                    throw;
                }

                var expectedError = File.ReadAllText(expectedErrorPath);
                if (expectedError != ex.Message)
                {
                    Assert.Fail($"Expected error: {expectedError}, but got: {ex.Message}");
                }

                return;
            }

            ExcelAssert.AreEqual(File.OpenRead(expectedPath), outputStream);
        }
    }

    public static class ExcelAssert
    {
        public static void AreEqual(Stream expectedStream, Stream actualStream)
        {
            using var expected = new ExcelPackage(expectedStream);
            using var actual = new ExcelPackage(actualStream);

            actual.Workbook.Worksheets.Count.ShouldBe(expected.Workbook.Worksheets.Count);

            for (var i = 0; i < expected.Workbook.Worksheets.Count; i++)
            {
                var wsExpected = expected.Workbook.Worksheets[i];
                var wsActual = actual.Workbook.Worksheets[i];

                wsActual.Dimension.Rows.ShouldBe(wsExpected.Dimension.Rows);
                wsActual.Dimension.Columns.ShouldBe(wsExpected.Dimension.Columns);

                if (wsExpected.Dimension == null) continue;

                for (var row = 1; row <= wsExpected.Dimension.End.Row; row++)
                {
                    for (var col = 1; col <= wsExpected.Dimension.End.Column; col++)
                    {
                        var cellExp = wsExpected.Cells[row, col];
                        var cellAct = wsActual.Cells[row, col];

                        cellAct.Text?.Trim().ShouldBe(cellExp.Text?.Trim());
                        cellAct.Style.Numberformat.Format.ShouldBe(cellExp.Style.Numberformat.Format);
                    }
                }
            }
        }
    }
}