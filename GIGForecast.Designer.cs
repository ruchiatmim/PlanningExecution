namespace Version3
{
    partial class GIGForecast
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel8 = new System.Windows.Forms.Panel();
            this.lblNumOfRec = new System.Windows.Forms.Label();
            this.btnProcess = new System.Windows.Forms.Button();
            this.btnUploadNewForecast = new System.Windows.Forms.Button();
            this.button14 = new System.Windows.Forms.Button();
            this.dgvNewforecast = new System.Windows.Forms.DataGridView();
            this.panel7 = new System.Windows.Forms.Panel();
            this.lblMappingTable = new System.Windows.Forms.Label();
            this.pnlPlanningItem = new System.Windows.Forms.Panel();
            this.label8 = new System.Windows.Forms.Label();
            this.spreadsheetListView1 = new System.Windows.Forms.ListView();
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.button9 = new System.Windows.Forms.Button();
            this.panel8.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvNewforecast)).BeginInit();
            this.panel7.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel8
            // 
            this.panel8.BackColor = System.Drawing.Color.BurlyWood;
            this.panel8.Controls.Add(this.lblNumOfRec);
            this.panel8.Controls.Add(this.btnProcess);
            this.panel8.Controls.Add(this.btnUploadNewForecast);
            this.panel8.Controls.Add(this.button14);
            this.panel8.Location = new System.Drawing.Point(1, 110);
            this.panel8.Name = "panel8";
            this.panel8.Size = new System.Drawing.Size(1920, 41);
            this.panel8.TabIndex = 30;
            // 
            // lblNumOfRec
            // 
            this.lblNumOfRec.AutoSize = true;
            this.lblNumOfRec.Location = new System.Drawing.Point(3, 25);
            this.lblNumOfRec.Name = "lblNumOfRec";
            this.lblNumOfRec.Size = new System.Drawing.Size(105, 13);
            this.lblNumOfRec.TabIndex = 37;
            this.lblNumOfRec.Text = "Number of Records :";
            // 
            // btnProcess
            // 
            this.btnProcess.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProcess.ForeColor = System.Drawing.Color.SaddleBrown;
            this.btnProcess.Location = new System.Drawing.Point(1082, 6);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(133, 22);
            this.btnProcess.TabIndex = 19;
            this.btnProcess.Text = "Process NSP Plan";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Visible = false;
            // 
            // btnUploadNewForecast
            // 
            this.btnUploadNewForecast.BackColor = System.Drawing.Color.Green;
            this.btnUploadNewForecast.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUploadNewForecast.ForeColor = System.Drawing.Color.White;
            this.btnUploadNewForecast.Location = new System.Drawing.Point(888, 7);
            this.btnUploadNewForecast.Name = "btnUploadNewForecast";
            this.btnUploadNewForecast.Size = new System.Drawing.Size(126, 31);
            this.btnUploadNewForecast.TabIndex = 17;
            this.btnUploadNewForecast.Text = "Upload File";
            this.btnUploadNewForecast.UseVisualStyleBackColor = false;
            this.btnUploadNewForecast.Click += new System.EventHandler(this.btnUploadNewForecast_Click);
            // 
            // button14
            // 
            this.button14.BackColor = System.Drawing.Color.DarkOrange;
            this.button14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button14.ForeColor = System.Drawing.Color.White;
            this.button14.Location = new System.Drawing.Point(686, 7);
            this.button14.Name = "button14";
            this.button14.Size = new System.Drawing.Size(122, 31);
            this.button14.TabIndex = 0;
            this.button14.Text = "Delete Row";
            this.button14.UseVisualStyleBackColor = false;
            // 
            // dgvNewforecast
            // 
            this.dgvNewforecast.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvNewforecast.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvNewforecast.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvNewforecast.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgvNewforecast.Location = new System.Drawing.Point(1, 152);
            this.dgvNewforecast.Name = "dgvNewforecast";
            this.dgvNewforecast.ReadOnly = true;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvNewforecast.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dgvNewforecast.Size = new System.Drawing.Size(1910, 801);
            this.dgvNewforecast.TabIndex = 29;
            this.dgvNewforecast.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvNewforecast_CellFormatting);
            // 
            // panel7
            // 
            this.panel7.BackColor = System.Drawing.Color.SaddleBrown;
            this.panel7.Controls.Add(this.lblMappingTable);
            this.panel7.Controls.Add(this.pnlPlanningItem);
            this.panel7.Controls.Add(this.label8);
            this.panel7.Controls.Add(this.spreadsheetListView1);
            this.panel7.Controls.Add(this.button9);
            this.panel7.Location = new System.Drawing.Point(1, 1);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(1920, 110);
            this.panel7.TabIndex = 28;
            // 
            // lblMappingTable
            // 
            this.lblMappingTable.AutoSize = true;
            this.lblMappingTable.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMappingTable.ForeColor = System.Drawing.Color.White;
            this.lblMappingTable.Location = new System.Drawing.Point(1219, 1);
            this.lblMappingTable.Name = "lblMappingTable";
            this.lblMappingTable.Size = new System.Drawing.Size(298, 16);
            this.lblMappingTable.TabIndex = 27;
            this.lblMappingTable.Text = "Planning Items are not in Mapping Table";
            this.lblMappingTable.Visible = false;
            // 
            // pnlPlanningItem
            // 
            this.pnlPlanningItem.AutoScroll = true;
            this.pnlPlanningItem.BackColor = System.Drawing.Color.PeachPuff;
            this.pnlPlanningItem.Cursor = System.Windows.Forms.Cursors.Default;
            this.pnlPlanningItem.ForeColor = System.Drawing.SystemColors.ControlText;
            this.pnlPlanningItem.Location = new System.Drawing.Point(1219, 20);
            this.pnlPlanningItem.Name = "pnlPlanningItem";
            this.pnlPlanningItem.Size = new System.Drawing.Size(550, 81);
            this.pnlPlanningItem.TabIndex = 26;
            this.pnlPlanningItem.Visible = false;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label8.Location = new System.Drawing.Point(6, 6);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(84, 13);
            this.label8.TabIndex = 24;
            this.label8.Text = "Doc Name  :";
            // 
            // spreadsheetListView1
            // 
            this.spreadsheetListView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader4});
            this.spreadsheetListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.spreadsheetListView1.HideSelection = false;
            this.spreadsheetListView1.LabelWrap = false;
            this.spreadsheetListView1.Location = new System.Drawing.Point(96, 3);
            this.spreadsheetListView1.MultiSelect = false;
            this.spreadsheetListView1.Name = "spreadsheetListView1";
            this.spreadsheetListView1.Size = new System.Drawing.Size(342, 97);
            this.spreadsheetListView1.TabIndex = 23;
            this.spreadsheetListView1.UseCompatibleStateImageBehavior = false;
            this.spreadsheetListView1.View = System.Windows.Forms.View.Details;
            this.spreadsheetListView1.SelectedIndexChanged += new System.EventHandler(this.spreadsheetListView1_SelectedIndexChanged);
            // 
            // columnHeader4
            // 
            this.columnHeader4.Width = 250;
            // 
            // button9
            // 
            this.button9.BackColor = System.Drawing.Color.BurlyWood;
            this.button9.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.button9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button9.ForeColor = System.Drawing.Color.Maroon;
            this.button9.Location = new System.Drawing.Point(444, 67);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(122, 31);
            this.button9.TabIndex = 22;
            this.button9.Text = "Download Grid";
            this.button9.UseVisualStyleBackColor = false;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // GIGForecast
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1868, 916);
            this.Controls.Add(this.panel8);
            this.Controls.Add(this.dgvNewforecast);
            this.Controls.Add(this.panel7);
            this.Name = "GIGForecast";
            this.Text = "GIGForecast";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.panel8.ResumeLayout(false);
            this.panel8.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvNewforecast)).EndInit();
            this.panel7.ResumeLayout(false);
            this.panel7.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel8;
        private System.Windows.Forms.Label lblNumOfRec;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Button btnUploadNewForecast;
        private System.Windows.Forms.Button button14;
        private System.Windows.Forms.DataGridView dgvNewforecast;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Label lblMappingTable;
        private System.Windows.Forms.Panel pnlPlanningItem;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ListView spreadsheetListView1;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.Button button9;
    }
}