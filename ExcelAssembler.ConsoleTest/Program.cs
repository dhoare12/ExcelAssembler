using System.Diagnostics;

namespace ExcelAssembler.ConsoleTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var templatePath = "C:\\Temp\\ExcelAssembler\\TestRepeatTemplate.xlsx";
            var xmlPath = "C:\\Temp\\ExcelAssembler\\RepeatTest.xml";

            var tmpPath = $"C:\\Temp\\ExcelAssembler\\Output{DateTime.UtcNow:yyyyMMddHHmmss}.xlsx";

            var excelAssembler = new ExcelAssembler(new ExcelAssemblerOptions
            {
                MissingXmlDataBehaviour = MissingXmlDataBehaviour.ShowPlaceholder
            });

            var stream = excelAssembler.ProcessTemplate(File.OpenRead(templatePath), File.ReadAllText(xmlPath));

            using (var fileStream = File.OpenWrite(tmpPath))
            {
                stream.CopyTo(fileStream);
            }

            var psi = new ProcessStartInfo(tmpPath)
            {
                UseShellExecute = true,
            };

            Process.Start(psi);
        }
    }
}
