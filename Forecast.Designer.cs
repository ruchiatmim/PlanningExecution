namespace Version3
{
    partial class Forecast
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
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.button6 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.listView2 = new System.Windows.Forms.ListView();
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.button7 = new System.Windows.Forms.Button();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panel8 = new System.Windows.Forms.Panel();
            this.lblNumOfRec = new System.Windows.Forms.Label();
            this.dgvNewforecast = new System.Windows.Forms.DataGridView();
            this.panel7 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.btnUploadNewForecast = new System.Windows.Forms.Button();
            this.lblMappingTable = new System.Windows.Forms.Label();
            this.pnlPlanningItem = new System.Windows.Forms.Panel();
            this.lslWorksheet1 = new System.Windows.Forms.ListBox();
            this.label8 = new System.Windows.Forms.Label();
            this.spreadsheetListView1 = new System.Windows.Forms.ListView();
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.button9 = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1.SuspendLayout();
            this.panel8.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvNewforecast)).BeginInit();
            this.panel7.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.ForeColor = System.Drawing.Color.SaddleBrown;
            this.button3.Location = new System.Drawing.Point(861, 7);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(120, 33);
            this.button3.TabIndex = 18;
            this.button3.Text = "Cancel";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Visible = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.ForeColor = System.Drawing.Color.SaddleBrown;
            this.button4.Location = new System.Drawing.Point(716, 7);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(126, 33);
            this.button4.TabIndex = 17;
            this.button4.Text = "Upload File";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button5.ForeColor = System.Drawing.Color.SaddleBrown;
            this.button5.Location = new System.Drawing.Point(554, 8);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(122, 31);
            this.button5.TabIndex = 0;
            this.button5.Text = "Delete Row";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label4.Location = new System.Drawing.Point(30, 6);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(161, 16);
            this.label4.TabIndex = 24;
            this.label4.Text = "Select Spreadsheet   :";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // listView1
            // 
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader2});
            this.listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.listView1.Location = new System.Drawing.Point(192, 6);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(320, 97);
            this.listView1.TabIndex = 23;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
            // 
            // columnHeader2
            // 
            this.columnHeader2.Width = 250;
            // 
            // button6
            // 
            this.button6.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.button6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button6.ForeColor = System.Drawing.Color.SaddleBrown;
            this.button6.Location = new System.Drawing.Point(1038, 72);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(122, 31);
            this.button6.TabIndex = 22;
            this.button6.Text = "Download Grid";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(680, 6);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(301, 82);
            this.listBox1.TabIndex = 21;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label5.Location = new System.Drawing.Point(518, 6);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(146, 16);
            this.label5.TabIndex = 20;
            this.label5.Text = "Select Worksheet   :";
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label6.Location = new System.Drawing.Point(30, 6);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(161, 16);
            this.label6.TabIndex = 24;
            this.label6.Text = "Select Spreadsheet   :";
            this.label6.Click += new System.EventHandler(this.label6_Click);
            // 
            // listView2
            // 
            this.listView2.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader3});
            this.listView2.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.listView2.Location = new System.Drawing.Point(192, 6);
            this.listView2.MultiSelect = false;
            this.listView2.Name = "listView2";
            this.listView2.Size = new System.Drawing.Size(320, 97);
            this.listView2.TabIndex = 23;
            this.listView2.UseCompatibleStateImageBehavior = false;
            this.listView2.View = System.Windows.Forms.View.Details;
            this.listView2.SelectedIndexChanged += new System.EventHandler(this.listView2_SelectedIndexChanged);
            // 
            // columnHeader3
            // 
            this.columnHeader3.Width = 250;
            // 
            // button7
            // 
            this.button7.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.button7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button7.ForeColor = System.Drawing.Color.SaddleBrown;
            this.button7.Location = new System.Drawing.Point(1038, 72);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(122, 31);
            this.button7.TabIndex = 22;
            this.button7.Text = "Download Grid";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // listBox2
            // 
            this.listBox2.FormattingEnabled = true;
            this.listBox2.Location = new System.Drawing.Point(680, 6);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(301, 82);
            this.listBox2.TabIndex = 21;
            this.listBox2.SelectedIndexChanged += new System.EventHandler(this.listBox2_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label7.Location = new System.Drawing.Point(518, 6);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(146, 16);
            this.label7.TabIndex = 20;
            this.label7.Text = "Select Worksheet   :";
            this.label7.Click += new System.EventHandler(this.label7_Click);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.panel8);
            this.tabPage1.Controls.Add(this.dgvNewforecast);
            this.tabPage1.Controls.Add(this.panel7);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1912, 1174);
            this.tabPage1.TabIndex = 3;
            this.tabPage1.Text = "Net Supply Plan";
            this.tabPage1.UseVisualStyleBackColor = true;
            this.tabPage1.Click += new System.EventHandler(this.tabPage1_Click);
            // 
            // panel8
            // 
            this.panel8.BackColor = System.Drawing.Color.BurlyWood;
            this.panel8.Controls.Add(this.lblNumOfRec);
            this.panel8.Location = new System.Drawing.Point(0, 112);
            this.panel8.Name = "panel8";
            this.panel8.Size = new System.Drawing.Size(1920, 22);
            this.panel8.TabIndex = 27;
            // 
            // lblNumOfRec
            // 
            this.lblNumOfRec.AutoSize = true;
            this.lblNumOfRec.Location = new System.Drawing.Point(3, 4);
            this.lblNumOfRec.Name = "lblNumOfRec";
            this.lblNumOfRec.Size = new System.Drawing.Size(126, 13);
            this.lblNumOfRec.TabIndex = 37;
            this.lblNumOfRec.Text = "Number of Records :";
            // 
            // dgvNewforecast
            // 
            this.dgvNewforecast.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvNewforecast.Location = new System.Drawing.Point(0, 134);
            this.dgvNewforecast.Name = "dgvNewforecast";
            this.dgvNewforecast.Size = new System.Drawing.Size(1910, 826);
            this.dgvNewforecast.TabIndex = 26;
            // 
            // panel7
            // 
            this.panel7.BackColor = System.Drawing.Color.SaddleBrown;
            this.panel7.Controls.Add(this.label2);
            this.panel7.Controls.Add(this.label1);
            this.panel7.Controls.Add(this.panel1);
            this.panel7.Controls.Add(this.btnUploadNewForecast);
            this.panel7.Controls.Add(this.lblMappingTable);
            this.panel7.Controls.Add(this.pnlPlanningItem);
            this.panel7.Controls.Add(this.lslWorksheet1);
            this.panel7.Controls.Add(this.label8);
            this.panel7.Controls.Add(this.spreadsheetListView1);
            this.panel7.Controls.Add(this.button9);
            this.panel7.Controls.Add(this.label9);
            this.panel7.Location = new System.Drawing.Point(0, 3);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(1920, 110);
            this.panel7.TabIndex = 25;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label2.Location = new System.Drawing.Point(1104, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(48, 13);
            this.label2.TabIndex = 31;
            this.label2.Text = "Step 2";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label1.Location = new System.Drawing.Point(940, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 13);
            this.label1.TabIndex = 30;
            this.label1.Text = "Step 1";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.radioButton3);
            this.panel1.Controls.Add(this.radioButton2);
            this.panel1.Controls.Add(this.radioButton1);
            this.panel1.Location = new System.Drawing.Point(698, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(205, 92);
            this.panel1.TabIndex = 29;
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButton3.ForeColor = System.Drawing.Color.White;
            this.radioButton3.Location = new System.Drawing.Point(12, 26);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(190, 17);
            this.radioButton3.TabIndex = 31;
            this.radioButton3.TabStop = true;
            this.radioButton3.Text = "Current Deployment Data";
            this.radioButton3.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButton2.ForeColor = System.Drawing.Color.White;
            this.radioButton2.Location = new System.Drawing.Point(12, 66);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(170, 17);
            this.radioButton2.TabIndex = 30;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "Last Weeks NSP Detail";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Checked = true;
            this.radioButton1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButton1.ForeColor = System.Drawing.Color.White;
            this.radioButton1.Location = new System.Drawing.Point(12, 3);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(145, 17);
            this.radioButton1.TabIndex = 29;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "Current NSP Detail";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // btnUploadNewForecast
            // 
            this.btnUploadNewForecast.BackColor = System.Drawing.Color.White;
            this.btnUploadNewForecast.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUploadNewForecast.ForeColor = System.Drawing.Color.SaddleBrown;
            this.btnUploadNewForecast.Location = new System.Drawing.Point(1064, 71);
            this.btnUploadNewForecast.Name = "btnUploadNewForecast";
            this.btnUploadNewForecast.Size = new System.Drawing.Size(126, 33);
            this.btnUploadNewForecast.TabIndex = 17;
            this.btnUploadNewForecast.Text = "Upload File";
            this.btnUploadNewForecast.UseVisualStyleBackColor = false;
            this.btnUploadNewForecast.Click += new System.EventHandler(this.btnUploadNewForecast_Click);
            // 
            // lblMappingTable
            // 
            this.lblMappingTable.AutoSize = true;
            this.lblMappingTable.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMappingTable.ForeColor = System.Drawing.Color.White;
            this.lblMappingTable.Location = new System.Drawing.Point(1193, 0);
            this.lblMappingTable.Name = "lblMappingTable";
            this.lblMappingTable.Size = new System.Drawing.Size(206, 16);
            this.lblMappingTable.TabIndex = 27;
            this.lblMappingTable.Text = "Planning Items not Mapped";
            this.lblMappingTable.Visible = false;
            // 
            // pnlPlanningItem
            // 
            this.pnlPlanningItem.AutoScroll = true;
            this.pnlPlanningItem.BackColor = System.Drawing.Color.PeachPuff;
            this.pnlPlanningItem.Cursor = System.Windows.Forms.Cursors.Default;
            this.pnlPlanningItem.ForeColor = System.Drawing.SystemColors.ControlText;
            this.pnlPlanningItem.Location = new System.Drawing.Point(1196, 20);
            this.pnlPlanningItem.Name = "pnlPlanningItem";
            this.pnlPlanningItem.Size = new System.Drawing.Size(401, 81);
            this.pnlPlanningItem.TabIndex = 26;
            this.pnlPlanningItem.Visible = false;
            // 
            // lslWorksheet1
            // 
            this.lslWorksheet1.FormattingEnabled = true;
            this.lslWorksheet1.Location = new System.Drawing.Point(351, 22);
            this.lslWorksheet1.Name = "lslWorksheet1";
            this.lslWorksheet1.Size = new System.Drawing.Size(322, 82);
            this.lslWorksheet1.TabIndex = 25;
            this.lslWorksheet1.SelectedIndexChanged += new System.EventHandler(this.lslWorksheet1_SelectedIndexChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label8.Location = new System.Drawing.Point(3, 6);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(80, 13);
            this.label8.TabIndex = 24;
            this.label8.Text = "Doc Name :";
            // 
            // spreadsheetListView1
            // 
            this.spreadsheetListView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader4});
            this.spreadsheetListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.spreadsheetListView1.HideSelection = false;
            this.spreadsheetListView1.LabelWrap = false;
            this.spreadsheetListView1.Location = new System.Drawing.Point(5, 22);
            this.spreadsheetListView1.MultiSelect = false;
            this.spreadsheetListView1.Name = "spreadsheetListView1";
            this.spreadsheetListView1.Size = new System.Drawing.Size(319, 79);
            this.spreadsheetListView1.TabIndex = 23;
            this.spreadsheetListView1.UseCompatibleStateImageBehavior = false;
            this.spreadsheetListView1.View = System.Windows.Forms.View.Details;
            this.spreadsheetListView1.DrawItem += new System.Windows.Forms.DrawListViewItemEventHandler(this.listView1_DrawItem);
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
            this.button9.Location = new System.Drawing.Point(909, 70);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(117, 34);
            this.button9.TabIndex = 22;
            this.button9.Text = "Download Grid";
            this.button9.UseVisualStyleBackColor = false;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label9.Location = new System.Drawing.Point(351, 6);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(93, 13);
            this.label9.TabIndex = 20;
            this.label9.Text = "Worksheet   :";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(2, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1920, 1200);
            this.tabControl1.TabIndex = 24;
            this.tabControl1.TabStop = false;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // Forecast
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1603, 753);
            this.Controls.Add(this.tabControl1);
            this.Name = "Forecast";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Net Supply Plan";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.tabPage1.ResumeLayout(false);
            this.panel8.ResumeLayout(false);
            this.panel8.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvNewforecast)).EndInit();
            this.panel7.ResumeLayout(false);
            this.panel7.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ListView listView2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Panel panel8;
        private System.Windows.Forms.Button btnUploadNewForecast;
        private System.Windows.Forms.DataGridView dgvNewforecast;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.ListBox lslWorksheet1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ListView spreadsheetListView1;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.Panel pnlPlanningItem;
        private System.Windows.Forms.Label lblMappingTable;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.Label lblNumOfRec;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;

    }
}

