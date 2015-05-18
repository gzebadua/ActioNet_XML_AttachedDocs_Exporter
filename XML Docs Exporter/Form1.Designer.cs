namespace XML_Docs_Exporter
{
    partial class IAAXMLExporter
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.rtb = new System.Windows.Forms.RichTextBox();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.inputFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // rtb
            // 
            this.rtb.Location = new System.Drawing.Point(12, 12);
            this.rtb.Name = "rtb";
            this.rtb.Size = new System.Drawing.Size(662, 214);
            this.rtb.TabIndex = 1;
            this.rtb.Text = "";
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(12, 232);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(115, 23);
            this.btnUpdate.TabIndex = 3;
            this.btnUpdate.Text = "Export";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnExit
            // 
            this.btnExit.Enabled = false;
            this.btnExit.Location = new System.Drawing.Point(599, 232);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 4;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // inputFolder
            // 
            this.inputFolder.ShowNewFolderButton = false;
            // 
            // IAAXMLExporter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(686, 269);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.rtb);
            this.Name = "IAAXMLExporter";
            this.Text = "IAA XML Exporter";
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.RichTextBox rtb;
        internal System.Windows.Forms.Button btnUpdate;
        internal System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.FolderBrowserDialog inputFolder;
    }
}

