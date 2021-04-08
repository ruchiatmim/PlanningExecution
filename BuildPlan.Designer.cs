namespace PlanningExecution
{
    partial class BuildPlan
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        /// 
        /// 

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
            Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
            Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
            Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tbpageUpload = new System.Windows.Forms.TabPage();
            this.dataGridView1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
            this.status = new System.Windows.Forms.StatusStrip();
            this.statusVersion = new System.Windows.Forms.ToolStripStatusLabel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.lblNumOfRow = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.btnUpload = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.spreadsheetListView = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnDownloadGrid = new System.Windows.Forms.Button();
            this.lstWorksheet = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.worksheetListView = new System.Windows.Forms.ListView();
            this.dataGridViewold = new System.Windows.Forms.DataGridView();
            this.tabControl1.SuspendLayout();
            this.tbpageUpload.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.status.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewold)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tbpageUpload);
            this.tabControl1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(2, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(2240, 1200);
            this.tabControl1.TabIndex = 24;
            this.tabControl1.TabStop = false;
            // 
            // tbpageUpload
            // 
            this.tbpageUpload.Controls.Add(this.dataGridView1);
            this.tbpageUpload.Controls.Add(this.status);
            this.tbpageUpload.Controls.Add(this.panel2);
            this.tbpageUpload.Controls.Add(this.panel1);
            this.tbpageUpload.Controls.Add(this.worksheetListView);
            this.tbpageUpload.Controls.Add(this.dataGridViewold);
            this.tbpageUpload.Location = new System.Drawing.Point(4, 22);
            this.tbpageUpload.Name = "tbpageUpload";
            this.tbpageUpload.Padding = new System.Windows.Forms.Padding(3);
            this.tbpageUpload.Size = new System.Drawing.Size(2232, 1174);
            this.tbpageUpload.TabIndex = 0;
            this.tbpageUpload.Text = "Upload Build Plan";
            this.tbpageUpload.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            appearance1.BackColor = System.Drawing.SystemColors.Window;
            appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption;
            this.dataGridView1.DisplayLayout.Appearance = appearance1;
            ultraGridBand1.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
            this.dataGridView1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
            this.dataGridView1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
            this.dataGridView1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.False;
            appearance2.BackColor = System.Drawing.SystemColors.ActiveBorder;
            appearance2.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            appearance2.BorderColor = System.Drawing.SystemColors.Window;
            this.dataGridView1.DisplayLayout.GroupByBox.Appearance = appearance2;
            appearance3.ForeColor = System.Drawing.SystemColors.GrayText;
            this.dataGridView1.DisplayLayout.GroupByBox.BandLabelAppearance = appearance3;
            this.dataGridView1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
            this.dataGridView1.DisplayLayout.GroupByBox.Hidden = true;
            appearance4.BackColor = System.Drawing.SystemColors.ControlLightLight;
            appearance4.BackColor2 = System.Drawing.SystemColors.Control;
            appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal;
            appearance4.ForeColor = System.Drawing.SystemColors.GrayText;
            this.dataGridView1.DisplayLayout.GroupByBox.PromptAppearance = appearance4;
            this.dataGridView1.DisplayLayout.MaxColScrollRegions = 1;
            this.dataGridView1.DisplayLayout.MaxRowScrollRegions = 1;
            appearance5.BackColor = System.Drawing.SystemColors.Window;
            appearance5.ForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridView1.DisplayLayout.Override.ActiveCellAppearance = appearance5;
            appearance6.BackColor = System.Drawing.SystemColors.Highlight;
            appearance6.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.dataGridView1.DisplayLayout.Override.ActiveRowAppearance = appearance6;
            this.dataGridView1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted;
            this.dataGridView1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted;
            appearance7.BackColor = System.Drawing.SystemColors.Window;
            this.dataGridView1.DisplayLayout.Override.CardAreaAppearance = appearance7;
            appearance8.BorderColor = System.Drawing.Color.Silver;
            appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter;
            this.dataGridView1.DisplayLayout.Override.CellAppearance = appearance8;
            this.dataGridView1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText;
            this.dataGridView1.DisplayLayout.Override.CellPadding = 0;
            appearance9.BackColor = System.Drawing.SystemColors.Control;
            appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element;
            appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal;
            appearance9.BorderColor = System.Drawing.SystemColors.Window;
            this.dataGridView1.DisplayLayout.Override.GroupByRowAppearance = appearance9;
            appearance10.TextHAlignAsString = "Left";
            this.dataGridView1.DisplayLayout.Override.HeaderAppearance = appearance10;
            this.dataGridView1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
            this.dataGridView1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand;
            appearance11.BackColor = System.Drawing.SystemColors.Window;
            appearance11.BorderColor = System.Drawing.Color.Silver;
            this.dataGridView1.DisplayLayout.Override.RowAppearance = appearance11;
            this.dataGridView1.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False;
            appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
            this.dataGridView1.DisplayLayout.Override.TemplateAddRowAppearance = appearance12;
            this.dataGridView1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
            this.dataGridView1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
            this.dataGridView1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
            this.dataGridView1.Location = new System.Drawing.Point(-28, 162);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1900, 780);
            this.dataGridView1.TabIndex = 27;
            this.dataGridView1.Text = "ultraGrid1";
            // 
            // status
            // 
            this.status.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusVersion});
            this.status.Location = new System.Drawing.Point(3, 1149);
            this.status.Name = "status";
            this.status.Padding = new System.Windows.Forms.Padding(1, 0, 16, 0);
            this.status.Size = new System.Drawing.Size(2226, 22);
            this.status.TabIndex = 26;
            this.status.Text = "statusStrip1";
            // 
            // statusVersion
            // 
            this.statusVersion.Name = "statusVersion";
            this.statusVersion.Size = new System.Drawing.Size(109, 17);
            this.statusVersion.Text = "toolStripStatusLabel1";
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.BurlyWood;
            this.panel2.Controls.Add(this.lblNumOfRow);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(this.button2);
            this.panel2.Controls.Add(this.btnUpload);
            this.panel2.Controls.Add(this.btnDelete);
            this.panel2.Location = new System.Drawing.Point(0, 109);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1900, 47);
            this.panel2.TabIndex = 25;
            // 
            // lblNumOfRow
            // 
            this.lblNumOfRow.AutoSize = true;
            this.lblNumOfRow.Location = new System.Drawing.Point(16, 25);
            this.lblNumOfRow.Name = "lblNumOfRow";
            this.lblNumOfRow.Size = new System.Drawing.Size(0, 13);
            this.lblNumOfRow.TabIndex = 27;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.BurlyWood;
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.button1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.Maroon;
            this.button1.Location = new System.Drawing.Point(1233, 8);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(142, 31);
            this.button1.TabIndex = 26;
            this.button1.Text = "Download Grid";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.ForeColor = System.Drawing.Color.Maroon;
            this.button2.Location = new System.Drawing.Point(1071, 7);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(140, 33);
            this.button2.TabIndex = 18;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Visible = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnUpload
            // 
            this.btnUpload.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpload.ForeColor = System.Drawing.Color.Maroon;
            this.btnUpload.Location = new System.Drawing.Point(838, 8);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(196, 33);
            this.btnUpload.TabIndex = 17;
            this.btnUpload.Text = "Process Build Plan";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDelete.ForeColor = System.Drawing.Color.SaddleBrown;
            this.btnDelete.Location = new System.Drawing.Point(646, 8);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(142, 31);
            this.btnDelete.TabIndex = 0;
            this.btnDelete.Text = "Delete Row";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Visible = false;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.SaddleBrown;
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.spreadsheetListView);
            this.panel1.Controls.Add(this.btnDownloadGrid);
            this.panel1.Controls.Add(this.lstWorksheet);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1900, 110);
            this.panel1.TabIndex = 24;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(1230, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 13);
            this.label2.TabIndex = 25;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label3.Location = new System.Drawing.Point(35, 6);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(149, 13);
            this.label3.TabIndex = 24;
            this.label3.Text = "Select Spreadsheet   :";
            // 
            // spreadsheetListView
            // 
            this.spreadsheetListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.spreadsheetListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.spreadsheetListView.Location = new System.Drawing.Point(224, 6);
            this.spreadsheetListView.MultiSelect = false;
            this.spreadsheetListView.Name = "spreadsheetListView";
            this.spreadsheetListView.Size = new System.Drawing.Size(373, 95);
            this.spreadsheetListView.TabIndex = 23;
            this.spreadsheetListView.UseCompatibleStateImageBehavior = false;
            this.spreadsheetListView.View = System.Windows.Forms.View.Details;
            this.spreadsheetListView.SelectedIndexChanged += new System.EventHandler(this.spreadsheetListView_SelectedIndexChanged);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Width = 250;
            // 
            // btnDownloadGrid
            // 
            this.btnDownloadGrid.BackColor = System.Drawing.Color.BurlyWood;
            this.btnDownloadGrid.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.btnDownloadGrid.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDownloadGrid.ForeColor = System.Drawing.Color.Maroon;
            this.btnDownloadGrid.Location = new System.Drawing.Point(621, 70);
            this.btnDownloadGrid.Name = "btnDownloadGrid";
            this.btnDownloadGrid.Size = new System.Drawing.Size(142, 31);
            this.btnDownloadGrid.TabIndex = 22;
            this.btnDownloadGrid.Text = "Download Grid";
            this.btnDownloadGrid.UseVisualStyleBackColor = false;
            this.btnDownloadGrid.Click += new System.EventHandler(this.button1_Click);
            // 
            // lstWorksheet
            // 
            this.lstWorksheet.FormattingEnabled = true;
            this.lstWorksheet.Location = new System.Drawing.Point(793, 6);
            this.lstWorksheet.Name = "lstWorksheet";
            this.lstWorksheet.Size = new System.Drawing.Size(373, 82);
            this.lstWorksheet.TabIndex = 21;
            this.lstWorksheet.Visible = false;
            this.lstWorksheet.SelectedIndexChanged += new System.EventHandler(this.lstWorksheet_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label1.Location = new System.Drawing.Point(604, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(137, 13);
            this.label1.TabIndex = 20;
            this.label1.Text = "Select Worksheet   :";
            this.label1.Visible = false;
            // 
            // worksheetListView
            // 
            this.worksheetListView.Location = new System.Drawing.Point(-346, -381);
            this.worksheetListView.MultiSelect = false;
            this.worksheetListView.Name = "worksheetListView";
            this.worksheetListView.ShowGroups = false;
            this.worksheetListView.Size = new System.Drawing.Size(33, 71);
            this.worksheetListView.TabIndex = 23;
            this.worksheetListView.UseCompatibleStateImageBehavior = false;
            this.worksheetListView.View = System.Windows.Forms.View.List;
            this.worksheetListView.Visible = false;
            // 
            // dataGridViewold
            // 
            this.dataGridViewold.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewold.Location = new System.Drawing.Point(861, 288);
            this.dataGridViewold.Name = "dataGridViewold";
            this.dataGridViewold.Size = new System.Drawing.Size(71, 15);
            this.dataGridViewold.TabIndex = 22;
            this.dataGridViewold.Visible = false;
            this.dataGridViewold.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // BuildPlan
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1985, 839);
            this.Controls.Add(this.tabControl1);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "BuildPlan";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Build Plan";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.tabControl1.ResumeLayout(false);
            this.tbpageUpload.ResumeLayout(false);
            this.tbpageUpload.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.status.ResumeLayout(false);
            this.status.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewold)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tbpageUpload;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnDownloadGrid;
        private System.Windows.Forms.ListBox lstWorksheet;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListView worksheetListView;
        private System.Windows.Forms.DataGridView dataGridViewold;
        private System.Windows.Forms.StatusStrip status;
        private System.Windows.Forms.ToolStripStatusLabel statusVersion;
        private System.Windows.Forms.ListView spreadsheetListView;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label lblNumOfRow;
        private Infragistics.Win.UltraWinGrid.UltraGrid dataGridView1;

    }
}

