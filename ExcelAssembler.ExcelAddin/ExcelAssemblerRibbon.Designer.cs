namespace ExcelAssembler.ExcelAddin
{
    partial class ExcelAssemblerRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExcelAssemblerRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnInsertContent = this.Factory.CreateRibbonButton();
            this.btnBeginRepeat = this.Factory.CreateRibbonButton();
            this.btnEndRepeat = this.Factory.CreateRibbonButton();
            this.btnTestXml = this.Factory.CreateRibbonButton();
            this.btnRetestLastFile = this.Factory.CreateRibbonButton();
            this.btnTogglePane = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnInsertContent);
            this.group1.Items.Add(this.btnBeginRepeat);
            this.group1.Items.Add(this.btnEndRepeat);
            this.group1.Items.Add(this.btnTestXml);
            this.group1.Items.Add(this.btnRetestLastFile);
            this.group1.Items.Add(this.btnTogglePane);
            this.group1.Label = "ExcelAssembler";
            this.group1.Name = "group1";
            // 
            // btnInsertContent
            // 
            this.btnInsertContent.Image = global::ExcelAssembler.ExcelAddin.Properties.Resources.variable;
            this.btnInsertContent.Label = "Insert Content";
            this.btnInsertContent.Name = "btnInsertContent";
            this.btnInsertContent.ShowImage = true;
            this.btnInsertContent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertContent_Click);
            // 
            // btnBeginRepeat
            // 
            this.btnBeginRepeat.Label = "Begin Repeat";
            this.btnBeginRepeat.Name = "btnBeginRepeat";
            this.btnBeginRepeat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBeginRepeat_Click);
            // 
            // btnEndRepeat
            // 
            this.btnEndRepeat.Label = "End Repeat";
            this.btnEndRepeat.Name = "btnEndRepeat";
            this.btnEndRepeat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEndRepeat_Click);
            // 
            // btnTestXml
            // 
            this.btnTestXml.Image = global::ExcelAssembler.ExcelAddin.Properties.Resources.fileCode;
            this.btnTestXml.Label = "Test with XML";
            this.btnTestXml.Name = "btnTestXml";
            this.btnTestXml.ShowImage = true;
            this.btnTestXml.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTestXml_Click);
            // 
            // btnRetestLastFile
            // 
            this.btnRetestLastFile.Label = "Retest Last File";
            this.btnRetestLastFile.Name = "btnRetestLastFile";
            this.btnRetestLastFile.Visible = false;
            this.btnRetestLastFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRetestLastFile_Click);
            // 
            // btnTogglePane
            // 
            this.btnTogglePane.Label = "Toggle Pane";
            this.btnTogglePane.Name = "btnTogglePane";
            this.btnTogglePane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTogglePane_Click);
            // 
            // ExcelAssemblerRibbon
            // 
            this.Name = "ExcelAssemblerRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExcelAssemblerRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTogglePane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertContent;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBeginRepeat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEndRepeat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTestXml;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRetestLastFile;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelAssemblerRibbon ExcelAssemblerRibbon
        {
            get { return this.GetRibbon<ExcelAssemblerRibbon>(); }
        }
    }
}
