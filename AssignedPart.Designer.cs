namespace Version3
{
    partial class AssignedPart
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtPartNum = new System.Windows.Forms.TextBox();
            this.lblPartNum = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.grdUnassigned = new System.Windows.Forms.DataGridView();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.btnExportConfirm = new System.Windows.Forms.Button();
            this.btnSaveConfType = new System.Windows.Forms.Button();
            this.btnSentForReview = new System.Windows.Forms.Button();
            this.btnLoadGrid = new System.Windows.Forms.Button();
            this.dgvAssignedConfirm = new System.Windows.Forms.DataGridView();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.panel4 = new System.Windows.Forms.Panel();
            this.btnExportUpdMIM = new System.Windows.Forms.Button();
            this.btnSaveINMIM = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.dgvAssignedUpdMIM = new System.Windows.Forms.DataGridView();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.panel5 = new System.Windows.Forms.Panel();
            this.btnMatchWindchill = new System.Windows.Forms.Button();
            this.chkMatchWindChill = new System.Windows.Forms.CheckBox();
            this.btnExportWindchill = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.dgvAssignedMatchWind = new System.Windows.Forms.DataGridView();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdUnassigned)).BeginInit();
            this.tabPage2.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAssignedConfirm)).BeginInit();
            this.tabPage3.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAssignedUpdMIM)).BeginInit();
            this.tabPage4.SuspendLayout();
            this.panel5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAssignedMatchWind)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(0, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1890, 1020);
            this.tabControl1.TabIndex = 3;
            this.tabControl1.Click += new System.EventHandler(this.tabControl1_Click);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.panel1);
            this.tabPage1.Controls.Add(this.grdUnassigned);
            this.tabPage1.Location = new System.Drawing.Point(4, 24);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1882, 992);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "New Unassigned Part";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.AutoSize = true;
            this.panel1.BackColor = System.Drawing.Color.SteelBlue;
            this.panel1.Controls.Add(this.txtPartNum);
            this.panel1.Controls.Add(this.lblPartNum);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.btnUpdate);
            this.panel1.Location = new System.Drawing.Point(6, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1800, 62);
            this.panel1.TabIndex = 4;
            // 
            // txtPartNum
            // 
            this.txtPartNum.Location = new System.Drawing.Point(117, 25);
            this.txtPartNum.Name = "txtPartNum";
            this.txtPartNum.Size = new System.Drawing.Size(183, 21);
            this.txtPartNum.TabIndex = 13;
            // 
            // lblPartNum
            // 
            this.lblPartNum.AutoSize = true;
            this.lblPartNum.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPartNum.ForeColor = System.Drawing.Color.White;
            this.lblPartNum.Location = new System.Drawing.Point(35, 28);
            this.lblPartNum.Name = "lblPartNum";
            this.lblPartNum.Size = new System.Drawing.Size(67, 15);
            this.lblPartNum.TabIndex = 12;
            this.lblPartNum.Text = "Part Num";
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.ForestGreen;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.ForeColor = System.Drawing.Color.White;
            this.button2.Location = new System.Drawing.Point(419, 21);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(118, 28);
            this.button2.TabIndex = 11;
            this.button2.Text = "Export To Excel";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.ForestGreen;
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.ForeColor = System.Drawing.Color.White;
            this.button3.Location = new System.Drawing.Point(318, 22);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(95, 27);
            this.button3.TabIndex = 8;
            this.button3.Text = "Load Grid";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.BackColor = System.Drawing.Color.ForestGreen;
            this.btnUpdate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpdate.ForeColor = System.Drawing.Color.White;
            this.btnUpdate.Location = new System.Drawing.Point(1212, 25);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(95, 27);
            this.btnUpdate.TabIndex = 4;
            this.btnUpdate.Text = "Save";
            this.btnUpdate.UseVisualStyleBackColor = false;
            this.btnUpdate.Visible = false;
            // 
            // grdUnassigned
            // 
            this.grdUnassigned.AllowUserToAddRows = false;
            this.grdUnassigned.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.grdUnassigned.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdUnassigned.Location = new System.Drawing.Point(6, 67);
            this.grdUnassigned.Name = "grdUnassigned";
            this.grdUnassigned.Size = new System.Drawing.Size(1800, 900);
            this.grdUnassigned.TabIndex = 3;
            this.grdUnassigned.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.grdUnassigned_CellContentClick);
            this.grdUnassigned.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.grdUnassigned_DataError);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.panel2);
            this.tabPage2.Controls.Add(this.dgvAssignedConfirm);
            this.tabPage2.Location = new System.Drawing.Point(4, 24);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1882, 992);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Sent For Review";
            this.tabPage2.UseVisualStyleBackColor = true;
            this.tabPage2.Click += new System.EventHandler(this.tabPage2_Click);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.SteelBlue;
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(this.btnExportConfirm);
            this.panel2.Controls.Add(this.btnSaveConfType);
            this.panel2.Controls.Add(this.btnLoadGrid);
            this.panel2.Controls.Add(this.btnSentForReview);
            this.panel2.Location = new System.Drawing.Point(0, 1);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1801, 44);
            this.panel2.TabIndex = 12;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Brown;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(1007, 11);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(131, 23);
            this.button1.TabIndex = 14;
            this.button1.Text = "Sync Parts";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnExportConfirm
            // 
            this.btnExportConfirm.BackColor = System.Drawing.Color.ForestGreen;
            this.btnExportConfirm.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportConfirm.ForeColor = System.Drawing.Color.White;
            this.btnExportConfirm.Location = new System.Drawing.Point(148, 9);
            this.btnExportConfirm.Name = "btnExportConfirm";
            this.btnExportConfirm.Size = new System.Drawing.Size(118, 26);
            this.btnExportConfirm.TabIndex = 13;
            this.btnExportConfirm.Text = "Export To Excel";
            this.btnExportConfirm.UseVisualStyleBackColor = false;
            this.btnExportConfirm.Click += new System.EventHandler(this.btnExportConfirm_Click);
            // 
            // btnSaveConfType
            // 
            this.btnSaveConfType.BackColor = System.Drawing.Color.Brown;
            this.btnSaveConfType.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveConfType.ForeColor = System.Drawing.Color.White;
            this.btnSaveConfType.Location = new System.Drawing.Point(864, 9);
            this.btnSaveConfType.Name = "btnSaveConfType";
            this.btnSaveConfType.Size = new System.Drawing.Size(137, 26);
            this.btnSaveConfType.TabIndex = 12;
            this.btnSaveConfType.Text = "Save";
            this.btnSaveConfType.UseVisualStyleBackColor = false;
            this.btnSaveConfType.Click += new System.EventHandler(this.btnSaveConfType_Click);
            // 
            // btnSentForReview
            // 
            this.btnSentForReview.BackColor = System.Drawing.Color.Brown;
            this.btnSentForReview.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSentForReview.ForeColor = System.Drawing.Color.White;
            this.btnSentForReview.Location = new System.Drawing.Point(1309, 9);
            this.btnSentForReview.Name = "btnSentForReview";
            this.btnSentForReview.Size = new System.Drawing.Size(127, 26);
            this.btnSentForReview.TabIndex = 11;
            this.btnSentForReview.Text = "Send For Review";
            this.btnSentForReview.UseVisualStyleBackColor = false;
            this.btnSentForReview.Visible = false;
            this.btnSentForReview.Click += new System.EventHandler(this.btnSentForReview_Click);
            // 
            // btnLoadGrid
            // 
            this.btnLoadGrid.BackColor = System.Drawing.Color.ForestGreen;
            this.btnLoadGrid.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLoadGrid.ForeColor = System.Drawing.Color.White;
            this.btnLoadGrid.Location = new System.Drawing.Point(6, 9);
            this.btnLoadGrid.Name = "btnLoadGrid";
            this.btnLoadGrid.Size = new System.Drawing.Size(10, 26);
            this.btnLoadGrid.TabIndex = 9;
            this.btnLoadGrid.Text = "Load Grid";
            this.btnLoadGrid.UseVisualStyleBackColor = false;
            this.btnLoadGrid.Visible = false;
            this.btnLoadGrid.Click += new System.EventHandler(this.btnLoadGrid_Click);
            // 
            // dgvAssignedConfirm
            // 
            this.dgvAssignedConfirm.AllowUserToAddRows = false;
            this.dgvAssignedConfirm.AllowUserToDeleteRows = false;
            this.dgvAssignedConfirm.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAssignedConfirm.Location = new System.Drawing.Point(1, 45);
            this.dgvAssignedConfirm.Name = "dgvAssignedConfirm";
            this.dgvAssignedConfirm.Size = new System.Drawing.Size(1800, 900);
            this.dgvAssignedConfirm.TabIndex = 6;
            this.dgvAssignedConfirm.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvAssigned_CellContentClick);
            this.dgvAssignedConfirm.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dgvAssigned_DataError);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.panel4);
            this.tabPage3.Controls.Add(this.dgvAssignedUpdMIM);
            this.tabPage3.Location = new System.Drawing.Point(4, 24);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(1882, 992);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Confirmed Supply Type";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.Color.SteelBlue;
            this.panel4.Controls.Add(this.btnExportUpdMIM);
            this.panel4.Controls.Add(this.btnSaveINMIM);
            this.panel4.Controls.Add(this.button6);
            this.panel4.Location = new System.Drawing.Point(3, 3);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1800, 44);
            this.panel4.TabIndex = 13;
            // 
            // btnExportUpdMIM
            // 
            this.btnExportUpdMIM.BackColor = System.Drawing.Color.ForestGreen;
            this.btnExportUpdMIM.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportUpdMIM.ForeColor = System.Drawing.Color.White;
            this.btnExportUpdMIM.Location = new System.Drawing.Point(127, 11);
            this.btnExportUpdMIM.Name = "btnExportUpdMIM";
            this.btnExportUpdMIM.Size = new System.Drawing.Size(118, 26);
            this.btnExportUpdMIM.TabIndex = 13;
            this.btnExportUpdMIM.Text = "Export To Excel";
            this.btnExportUpdMIM.UseVisualStyleBackColor = false;
            this.btnExportUpdMIM.Click += new System.EventHandler(this.btnExportUpdMIM_Click);
            // 
            // btnSaveINMIM
            // 
            this.btnSaveINMIM.BackColor = System.Drawing.Color.Brown;
            this.btnSaveINMIM.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveINMIM.ForeColor = System.Drawing.Color.White;
            this.btnSaveINMIM.Location = new System.Drawing.Point(856, 9);
            this.btnSaveINMIM.Name = "btnSaveINMIM";
            this.btnSaveINMIM.Size = new System.Drawing.Size(125, 26);
            this.btnSaveINMIM.TabIndex = 12;
            this.btnSaveINMIM.Text = "Save";
            this.btnSaveINMIM.UseVisualStyleBackColor = false;
            this.btnSaveINMIM.Click += new System.EventHandler(this.btnSaveINMIM_Click);
            // 
            // button6
            // 
            this.button6.BackColor = System.Drawing.Color.ForestGreen;
            this.button6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button6.ForeColor = System.Drawing.Color.White;
            this.button6.Location = new System.Drawing.Point(12, 11);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(21, 26);
            this.button6.TabIndex = 9;
            this.button6.Text = "Load Grid";
            this.button6.UseVisualStyleBackColor = false;
            this.button6.Visible = false;
            // 
            // dgvAssignedUpdMIM
            // 
            this.dgvAssignedUpdMIM.AllowUserToAddRows = false;
            this.dgvAssignedUpdMIM.AllowUserToDeleteRows = false;
            this.dgvAssignedUpdMIM.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAssignedUpdMIM.Location = new System.Drawing.Point(0, 46);
            this.dgvAssignedUpdMIM.Name = "dgvAssignedUpdMIM";
            this.dgvAssignedUpdMIM.Size = new System.Drawing.Size(1800, 900);
            this.dgvAssignedUpdMIM.TabIndex = 7;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.panel5);
            this.tabPage4.Controls.Add(this.dgvAssignedMatchWind);
            this.tabPage4.Location = new System.Drawing.Point(4, 24);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(1882, 992);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Updated In MIM";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.Color.SteelBlue;
            this.panel5.Controls.Add(this.btnMatchWindchill);
            this.panel5.Controls.Add(this.chkMatchWindChill);
            this.panel5.Controls.Add(this.btnExportWindchill);
            this.panel5.Controls.Add(this.button9);
            this.panel5.Location = new System.Drawing.Point(3, 3);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(1800, 44);
            this.panel5.TabIndex = 15;
            // 
            // btnMatchWindchill
            // 
            this.btnMatchWindchill.BackColor = System.Drawing.Color.Brown;
            this.btnMatchWindchill.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMatchWindchill.ForeColor = System.Drawing.Color.White;
            this.btnMatchWindchill.Location = new System.Drawing.Point(703, 10);
            this.btnMatchWindchill.Name = "btnMatchWindchill";
            this.btnMatchWindchill.Size = new System.Drawing.Size(118, 26);
            this.btnMatchWindchill.TabIndex = 16;
            this.btnMatchWindchill.Text = "Save";
            this.btnMatchWindchill.UseVisualStyleBackColor = false;
            this.btnMatchWindchill.Click += new System.EventHandler(this.btnMatchWindchill_Click);
            // 
            // chkMatchWindChill
            // 
            this.chkMatchWindChill.AutoSize = true;
            this.chkMatchWindChill.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkMatchWindChill.ForeColor = System.Drawing.Color.White;
            this.chkMatchWindChill.Location = new System.Drawing.Point(537, 15);
            this.chkMatchWindChill.Name = "chkMatchWindChill";
            this.chkMatchWindChill.Size = new System.Drawing.Size(128, 19);
            this.chkMatchWindChill.TabIndex = 15;
            this.chkMatchWindChill.Text = "Match Windchill";
            this.chkMatchWindChill.UseVisualStyleBackColor = true;
            // 
            // btnExportWindchill
            // 
            this.btnExportWindchill.BackColor = System.Drawing.Color.ForestGreen;
            this.btnExportWindchill.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportWindchill.ForeColor = System.Drawing.Color.White;
            this.btnExportWindchill.Location = new System.Drawing.Point(105, 10);
            this.btnExportWindchill.Name = "btnExportWindchill";
            this.btnExportWindchill.Size = new System.Drawing.Size(118, 26);
            this.btnExportWindchill.TabIndex = 14;
            this.btnExportWindchill.Text = "Export To Excel";
            this.btnExportWindchill.UseVisualStyleBackColor = false;
            this.btnExportWindchill.Click += new System.EventHandler(this.btnExportWindchill_Click);
            // 
            // button9
            // 
            this.button9.BackColor = System.Drawing.Color.ForestGreen;
            this.button9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button9.ForeColor = System.Drawing.Color.White;
            this.button9.Location = new System.Drawing.Point(12, 10);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(27, 26);
            this.button9.TabIndex = 9;
            this.button9.Text = "Load Grid";
            this.button9.UseVisualStyleBackColor = false;
            this.button9.Visible = false;
            // 
            // dgvAssignedMatchWind
            // 
            this.dgvAssignedMatchWind.AllowUserToAddRows = false;
            this.dgvAssignedMatchWind.AllowUserToDeleteRows = false;
            this.dgvAssignedMatchWind.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAssignedMatchWind.Location = new System.Drawing.Point(3, 45);
            this.dgvAssignedMatchWind.Name = "dgvAssignedMatchWind";
            this.dgvAssignedMatchWind.Size = new System.Drawing.Size(1800, 900);
            this.dgvAssignedMatchWind.TabIndex = 14;
            // 
            // AssignedPart
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(1882, 873);
            this.Controls.Add(this.tabControl1);
            this.Name = "AssignedPart";
            this.Text = "Unassigned Part";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdUnassigned)).EndInit();
            this.tabPage2.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvAssignedConfirm)).EndInit();
            this.tabPage3.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvAssignedUpdMIM)).EndInit();
            this.tabPage4.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAssignedMatchWind)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.DataGridView grdUnassigned;
        private System.Windows.Forms.DataGridView dgvAssignedConfirm;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btnLoadGrid;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox txtPartNum;
        private System.Windows.Forms.Label lblPartNum;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.DataGridView dgvAssignedUpdMIM;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Button btnSentForReview;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.DataGridView dgvAssignedMatchWind;
        private System.Windows.Forms.Button btnExportConfirm;
        private System.Windows.Forms.Button btnSaveConfType;
        private System.Windows.Forms.Button btnExportUpdMIM;
        private System.Windows.Forms.Button btnSaveINMIM;
        private System.Windows.Forms.Button btnMatchWindchill;
        private System.Windows.Forms.CheckBox chkMatchWindChill;
        private System.Windows.Forms.Button btnExportWindchill;
        private System.Windows.Forms.Button button1;

    }
}