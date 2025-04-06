using NUnit.Framework;

namespace ExcelAssembler.Tests;

public class ExcelAssemblerTests
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
            
        var excelAssembler = new ExcelAssembler(new ExcelAssemblerOptions
        {
            MissingXmlDataBehaviour = MissingXmlDataBehaviour.ThrowException
        });

        Stream outputStream;
        try
        {
            outputStream = excelAssembler.ProcessTemplate(File.OpenRead(templatePath), File.ReadAllText(xmlPath));
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