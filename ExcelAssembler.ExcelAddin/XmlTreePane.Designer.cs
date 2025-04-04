namespace ExcelAssembler.ExcelAddin
{
    partial class XmlTreePane
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.components = new System.ComponentModel.Container();
            this.treeTokens = new System.Windows.Forms.TreeView();
            this.btnLoadXmlFile = new System.Windows.Forms.Button();
            this.contextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menuItemInsertContent = new System.Windows.Forms.ToolStripMenuItem();
            this.menuItemInsertRepeat = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeTokens
            // 
            this.treeTokens.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.treeTokens.Location = new System.Drawing.Point(3, 92);
            this.treeTokens.Name = "treeTokens";
            this.treeTokens.Size = new System.Drawing.Size(896, 1167);
            this.treeTokens.TabIndex = 0;
            this.treeTokens.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeTokens_NodeMouseClick);
            this.treeTokens.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeTokens_NodeMouseDoubleClick);
            // 
            // btnLoadXmlFile
            // 
            this.btnLoadXmlFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnLoadXmlFile.Location = new System.Drawing.Point(3, 3);
            this.btnLoadXmlFile.Name = "btnLoadXmlFile";
            this.btnLoadXmlFile.Size = new System.Drawing.Size(896, 83);
            this.btnLoadXmlFile.TabIndex = 1;
            this.btnLoadXmlFile.Text = "Load XML File";
            this.btnLoadXmlFile.UseVisualStyleBackColor = true;
            this.btnLoadXmlFile.Click += new System.EventHandler(this.btnLoadXmlFile_Click);
            // 
            // contextMenu
            // 
            this.contextMenu.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.contextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuItemInsertContent,
            this.menuItemInsertRepeat});
            this.contextMenu.Name = "contextMenu";
            this.contextMenu.Size = new System.Drawing.Size(241, 80);
            // 
            // menuItemInsertContent
            // 
            this.menuItemInsertContent.Name = "menuItemInsertContent";
            this.menuItemInsertContent.Size = new System.Drawing.Size(240, 38);
            this.menuItemInsertContent.Text = "Insert Content";
            this.menuItemInsertContent.Click += new System.EventHandler(this.menuItemInsertContent_Click);
            // 
            // menuItemInsertRepeat
            // 
            this.menuItemInsertRepeat.Name = "menuItemInsertRepeat";
            this.menuItemInsertRepeat.Size = new System.Drawing.Size(240, 38);
            this.menuItemInsertRepeat.Text = "Insert Repeat";
            this.menuItemInsertRepeat.Click += new System.EventHandler(this.menuItemInsertRepeat_Click);
            // 
            // XmlTreePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnLoadXmlFile);
            this.Controls.Add(this.treeTokens);
            this.Name = "XmlTreePane";
            this.Size = new System.Drawing.Size(902, 1262);
            this.contextMenu.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeTokens;
        private System.Windows.Forms.Button btnLoadXmlFile;
        private System.Windows.Forms.ContextMenuStrip contextMenu;
        private System.Windows.Forms.ToolStripMenuItem menuItemInsertContent;
        private System.Windows.Forms.ToolStripMenuItem menuItemInsertRepeat;
    }
}
