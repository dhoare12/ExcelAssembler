using Microsoft.Office.Tools.Ribbon;
using System.Linq;

namespace ExcelAssembler.ExcelAddin
{
    public partial class ExcelAssemblerRibbon
    {
        private void ExcelAssemblerRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            var pane = Globals.ThisAddIn.CustomTaskPanes
                .FirstOrDefault(p => p.Control is XmlTreePane);

            if (pane != null)
            {
                pane.Visible = !pane.Visible;
            }
        }
    }
}
