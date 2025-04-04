using System.Diagnostics;

namespace ExcelAssembler
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var templatePath = "C:\\Temp\\ExcelAssembler\\TestRepeatTemplate.xlsx";
            var xmlPath = "C:\\Temp\\ExcelAssembler\\RepeatTest.xml";

            var tmpPath = $"C:\\Temp\\ExcelAssembler\\Output{DateTime.UtcNow:yyyyMMddHHmmss}.xlsx";

            var assembler = new ExcelAssembler();
            assembler.ProcessTemplate(templatePath, xmlPath, tmpPath);

            var psi = new ProcessStartInfo(tmpPath)
            {
                UseShellExecute = true,
            };

            Process.Start(psi);

        }
    }
}
