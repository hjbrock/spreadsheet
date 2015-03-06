namespace SS
{
    partial class SpreadsheetForm
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
            this.spreadsheetPanel = new SS.SpreadsheetPanel();
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.newToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editPanel = new System.Windows.Forms.Panel();
            this.TxtCellValue = new System.Windows.Forms.TextBox();
            this.TxtCellName = new System.Windows.Forms.TextBox();
            this.LblCellName = new System.Windows.Forms.Label();
            this.TxtCellContent = new System.Windows.Forms.TextBox();
            this.LblCellValue = new System.Windows.Forms.Label();
            this.LblEditCell = new System.Windows.Forms.Label();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpContentsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpProvider = new System.Windows.Forms.HelpProvider();
            this.menuStrip.SuspendLayout();
            this.editPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // spreadsheetPanel
            // 
            this.spreadsheetPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.spreadsheetPanel.Location = new System.Drawing.Point(0, 54);
            this.spreadsheetPanel.Name = "spreadsheetPanel";
            this.spreadsheetPanel.Size = new System.Drawing.Size(612, 410);
            this.spreadsheetPanel.TabIndex = 0;
            // 
            // menuStrip
            // 
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new System.Drawing.Size(612, 24);
            this.menuStrip.TabIndex = 1;
            this.menuStrip.Text = "menuStrip";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newToolStripMenuItem,
            this.openToolStripMenuItem,
            this.saveToolStripMenuItem,
            this.saveAsToolStripMenuItem,
            this.closeToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // newToolStripMenuItem
            // 
            this.newToolStripMenuItem.Name = "newToolStripMenuItem";
            this.newToolStripMenuItem.Size = new System.Drawing.Size(121, 22);
            this.newToolStripMenuItem.Text = "New";
            this.newToolStripMenuItem.Click += new System.EventHandler(this.newToolStripMenuItem_Click);
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(121, 22);
            this.openToolStripMenuItem.Text = "Open...";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(121, 22);
            this.saveToolStripMenuItem.Text = "Save";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // saveAsToolStripMenuItem
            // 
            this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(121, 22);
            this.saveAsToolStripMenuItem.Text = "Save as...";
            this.saveAsToolStripMenuItem.Click += new System.EventHandler(this.saveAsToolStripMenuItem_Click);
            // 
            // closeToolStripMenuItem
            // 
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            this.closeToolStripMenuItem.Size = new System.Drawing.Size(121, 22);
            this.closeToolStripMenuItem.Text = "Close";
            this.closeToolStripMenuItem.Click += new System.EventHandler(this.closeToolStripMenuItem_Click);
            // 
            // editPanel
            // 
            this.editPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.editPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.editPanel.Controls.Add(this.TxtCellValue);
            this.editPanel.Controls.Add(this.TxtCellName);
            this.editPanel.Controls.Add(this.LblCellName);
            this.editPanel.Controls.Add(this.TxtCellContent);
            this.editPanel.Controls.Add(this.LblCellValue);
            this.editPanel.Controls.Add(this.LblEditCell);
            this.editPanel.Location = new System.Drawing.Point(0, 27);
            this.editPanel.Name = "editPanel";
            this.editPanel.Size = new System.Drawing.Size(612, 21);
            this.editPanel.TabIndex = 2;
            // 
            // TxtCellValue
            // 
            this.TxtCellValue.Location = new System.Drawing.Point(116, 1);
            this.TxtCellValue.Name = "TxtCellValue";
            this.TxtCellValue.ReadOnly = true;
            this.TxtCellValue.Size = new System.Drawing.Size(95, 20);
            this.TxtCellValue.TabIndex = 5;
            // 
            // TxtCellName
            // 
            this.TxtCellName.Location = new System.Drawing.Point(42, 1);
            this.TxtCellName.Name = "TxtCellName";
            this.TxtCellName.ReadOnly = true;
            this.TxtCellName.Size = new System.Drawing.Size(28, 20);
            this.TxtCellName.TabIndex = 4;
            // 
            // LblCellName
            // 
            this.LblCellName.AutoSize = true;
            this.LblCellName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblCellName.Location = new System.Drawing.Point(3, 5);
            this.LblCellName.Name = "LblCellName";
            this.LblCellName.Size = new System.Drawing.Size(32, 13);
            this.LblCellName.TabIndex = 3;
            this.LblCellName.Text = "Cell:";
            // 
            // TxtCellContent
            // 
            this.TxtCellContent.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.TxtCellContent.Location = new System.Drawing.Point(269, 1);
            this.TxtCellContent.Name = "TxtCellContent";
            this.TxtCellContent.Size = new System.Drawing.Size(331, 20);
            this.TxtCellContent.TabIndex = 1;
            this.TxtCellContent.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxtCellContent_KeyDown);
            // 
            // LblCellValue
            // 
            this.LblCellValue.AutoSize = true;
            this.LblCellValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblCellValue.Location = new System.Drawing.Point(76, 5);
            this.LblCellValue.Name = "LblCellValue";
            this.LblCellValue.Size = new System.Drawing.Size(43, 13);
            this.LblCellValue.TabIndex = 2;
            this.LblCellValue.Text = "Value:";
            // 
            // LblEditCell
            // 
            this.LblEditCell.AutoSize = true;
            this.LblEditCell.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblEditCell.Location = new System.Drawing.Point(217, 5);
            this.LblEditCell.Name = "LblEditCell";
            this.LblEditCell.Size = new System.Drawing.Size(55, 13);
            this.LblEditCell.TabIndex = 0;
            this.LblEditCell.Text = "Content:";
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.helpContentsToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // helpContentsToolStripMenuItem
            // 
            this.helpContentsToolStripMenuItem.Name = "helpContentsToolStripMenuItem";
            this.helpContentsToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.helpContentsToolStripMenuItem.Text = "Help Contents";
            this.helpContentsToolStripMenuItem.Click += new System.EventHandler(this.helpContentsToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // helpProvider
            // 
            this.helpProvider.HelpNamespace = "SpreadsheetGUI.Properties.Resources.ps6_help.pdf";
            // 
            // SpreadsheetForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(612, 461);
            this.Controls.Add(this.editPanel);
            this.Controls.Add(this.spreadsheetPanel);
            this.Controls.Add(this.menuStrip);
            this.MainMenuStrip = this.menuStrip;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(627, 499);
            this.Name = "SpreadsheetForm";
            this.Text = "New Spreadsheet";
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.editPanel.ResumeLayout(false);
            this.editPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private SpreadsheetPanel spreadsheetPanel;
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem newToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.Panel editPanel;
        private System.Windows.Forms.Label LblCellValue;
        private System.Windows.Forms.TextBox TxtCellContent;
        private System.Windows.Forms.Label LblEditCell;
        private System.Windows.Forms.TextBox TxtCellValue;
        private System.Windows.Forms.TextBox TxtCellName;
        private System.Windows.Forms.Label LblCellName;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpContentsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.HelpProvider helpProvider;



    }
}

