using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.Office.Interop.Excel;

namespace ExcelAssembler.ExcelAddin
{
    public partial class ExcelAssemblerRibbon
    {
        private string _xmlPath;

        private void ExcelAssemblerRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void btnTestXml_Click(object sender, RibbonControlEventArgs e)
        {
            // Ask user to pick XML file
            var ofd = new OpenFileDialog
            {
                Filter = "XML Files (*.xml)|*.xml",
                Title = "Select XML Data File"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            btnRetestLastFile.Visible = true;

            _xmlPath = ofd.FileName;

            TestXmlFile();
        }

        private void TestXmlFile()
        {
            // Save current Word doc to a temp path
            var tempTemplatePath = Path.GetTempFileName() + ".xlsx";
            var doc = Globals.ThisAddIn.Application.ActiveWorkbook;

            doc.Save();
            File.Copy(doc.FullName, tempTemplatePath);

            // Load the XML
            var data = File.ReadAllText(_xmlPath);

            var templateStream = new MemoryStream();
            using (var fileStream = new FileStream(tempTemplatePath, FileMode.Open, FileAccess.Read))
            {
                fileStream.CopyTo(templateStream);
            }

            var resultPath = Path.Combine(Path.GetTempPath(), "Generated_" + Path.GetFileName(tempTemplatePath));

            var output = ExcelAssembler.ProcessTemplate(templateStream, data, suppressMissingXml: true);

            using (var outputFile = new FileStream(resultPath, FileMode.Create, FileAccess.Write))
            {
                output.CopyTo(outputFile);
            }

            // Open in Word
            var wordApp = new Microsoft.Office.Interop.Excel.Application();
            wordApp.Visible = true;
            wordApp.Workbooks.Open(resultPath);
        }

        private void btnRetestLastFile_Click(object sender, RibbonControlEventArgs e)
        {
            TestXmlFile();
        }

        private void btnTogglePane_Click(object sender, RibbonControlEventArgs e)
        {
            var pane = Globals.ThisAddIn.CustomTaskPanes
                .FirstOrDefault(p => p.Control is XmlTreePane);

            if (pane != null)
            {
                pane.Visible = !pane.Visible;
            }
        }

        private void btnInsertContent_Click(object sender, RibbonControlEventArgs e)
        {
            var input = Interaction.InputBox("Insert OpenXml content", "Enter the XPath for the data you'd like to be displayed below in (the cell below) this cell");
            var xpath = string.IsNullOrWhiteSpace(input) ? null : input.Trim();
            var activeCell = Globals.ThisAddIn.Application.ActiveCell;
            if (activeCell == null)
            {
                return;
            }
            activeCell.Value2 = $"<Content Select=\"{xpath}\" />";
        }

        private void btnBeginRepeat_Click(object sender, RibbonControlEventArgs e)
        {
            var input = Interaction.InputBox("Insert OpenXml content", "Enter the XPath for the first row of the data you'd like to loop through (eg ./Policy/Coverages/Coverages, not ./Policy/Coverages)");
            var xpath = string.IsNullOrWhiteSpace(input) ? null : input.Trim();
            var activeCell = Globals.ThisAddIn.Application.ActiveCell;
            if (activeCell == null)
            {
                return;
            }
            activeCell.Value2 = $"<Repeat Select=\"{xpath}\" />";
        }

        private void btnEndRepeat_Click(object sender, RibbonControlEventArgs e)
        {
            var activeCell = Globals.ThisAddIn.Application.ActiveCell;
            if (activeCell == null)
            {
                return;
            }
            activeCell.Value2 = "<EndRepeat />";
        }
    }
}
