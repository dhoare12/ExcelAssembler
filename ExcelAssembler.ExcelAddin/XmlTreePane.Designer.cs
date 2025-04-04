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
            this.treeTokens = new System.Windows.Forms.TreeView();
            this.btnLoadXmlFile = new System.Windows.Forms.Button();
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
            // XmlTreePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnLoadXmlFile);
            this.Controls.Add(this.treeTokens);
            this.Name = "XmlTreePane";
            this.Size = new System.Drawing.Size(902, 1262);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeTokens;
        private System.Windows.Forms.Button btnLoadXmlFile;
    }
}
