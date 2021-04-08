namespace Version3
{
    partial class NPI_MRP
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
            this.Panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.txtBuildID = new System.Windows.Forms.TextBox();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.lblProjectName = new System.Windows.Forms.Label();
            this.txtProjectName = new System.Windows.Forms.TextBox();
            this.cmbEUShipMeth = new System.Windows.Forms.ComboBox();
            this.cmbEUFOB = new System.Windows.Forms.ComboBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnUpload = new System.Windows.Forms.Button();
            this.btnDownloadGrid = new System.Windows.Forms.Button();
            this.Panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // Panel1
            // 
            this.Panel1.BackColor = System.Drawing.Color.SaddleBrown;
            this.Panel1.Controls.Add(this.label1);
            this.Panel1.Controls.Add(this.txtBuildID);
            this.Panel1.Controls.Add(this.radioButton2);
            this.Panel1.Controls.Add(this.radioButton1);
            this.Panel1.Controls.Add(this.lblProjectName);
            this.Panel1.Controls.Add(this.txtProjectName);
            this.Panel1.Controls.Add(this.cmbEUShipMeth);
            this.Panel1.Controls.Add(this.cmbEUFOB);
            this.Panel1.Location = new System.Drawing.Point(0, 0);
            this.Panel1.Name = "Panel1";
            this.Panel1.Size = new System.Drawing.Size(2240, 92);
            this.Panel1.TabIndex = 31;
            this.Panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.Panel1_Paint);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(165, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 16);
            this.label1.TabIndex = 51;
            this.label1.Text = "Build ID :";
            // 
            // txtBuildID
            // 
            this.txtBuildID.Location = new System.Drawing.Point(267, 47);
            this.txtBuildID.Name = "txtBuildID";
            this.txtBuildID.Size = new System.Drawing.Size(694, 20);
            this.txtBuildID.TabIndex = 50;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(133, 22);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(14, 13);
            this.radioButton2.TabIndex = 49;
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.Visible = false;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Checked = true;
            this.radioButton1.ForeColor = System.Drawing.Color.White;
            this.radioButton1.Location = new System.Drawing.Point(133, 51);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(14, 13);
            this.radioButton1.TabIndex = 48;
            this.radioButton1.TabStop = true;
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.Visible = false;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // lblProjectName
            // 
            this.lblProjectName.AutoSize = true;
            this.lblProjectName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblProjectName.ForeColor = System.Drawing.Color.White;
            this.lblProjectName.Location = new System.Drawing.Point(165, 20);
            this.lblProjectName.Name = "lblProjectName";
            this.lblProjectName.Size = new System.Drawing.Size(96, 16);
            this.lblProjectName.TabIndex = 47;
            this.lblProjectName.Text = "Project Name :";
            this.lblProjectName.Visible = false;
            // 
            // txtProjectName
            // 
            this.txtProjectName.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtProjectName.Location = new System.Drawing.Point(267, 19);
            this.txtProjectName.Name = "txtProjectName";
            this.txtProjectName.Size = new System.Drawing.Size(694, 21);
            this.txtProjectName.TabIndex = 46;
            this.txtProjectName.Visible = false;
            // 
            // cmbEUShipMeth
            // 
            this.cmbEUShipMeth.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbEUShipMeth.FormattingEnabled = true;
            this.cmbEUShipMeth.Location = new System.Drawing.Point(1517, 43);
            this.cmbEUShipMeth.Name = "cmbEUShipMeth";
            this.cmbEUShipMeth.Size = new System.Drawing.Size(282, 22);
            this.cmbEUShipMeth.TabIndex = 35;
            // 
            // cmbEUFOB
            // 
            this.cmbEUFOB.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbEUFOB.FormattingEnabled = true;
            this.cmbEUFOB.Location = new System.Drawing.Point(1517, 6);
            this.cmbEUFOB.Name = "cmbEUFOB";
            this.cmbEUFOB.Size = new System.Drawing.Size(282, 22);
            this.cmbEUFOB.TabIndex = 34;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.BurlyWood;
            this.panel2.Controls.Add(this.btnUpload);
            this.panel2.Controls.Add(this.btnDownloadGrid);
            this.panel2.Location = new System.Drawing.Point(0, 92);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(2240, 47);
            this.panel2.TabIndex = 32;
            // 
            // btnUpload
            // 
            this.btnUpload.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpload.ForeColor = System.Drawing.Color.SaddleBrown;
            this.btnUpload.Location = new System.Drawing.Point(880, 6);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(143, 33);
            this.btnUpload.TabIndex = 17;
            this.btnUpload.Text = "Update Shared Doc";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // btnDownloadGrid
            // 
            this.btnDownloadGrid.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDownloadGrid.ForeColor = System.Drawing.Color.SaddleBrown;
            this.btnDownloadGrid.Location = new System.Drawing.Point(667, 8);
            this.btnDownloadGrid.Name = "btnDownloadGrid";
            this.btnDownloadGrid.Size = new System.Drawing.Size(142, 31);
            this.btnDownloadGrid.TabIndex = 22;
            this.btnDownloadGrid.Text = "Get data in Excel";
            this.btnDownloadGrid.UseVisualStyleBackColor = true;
            this.btnDownloadGrid.Click += new System.EventHandler(this.btnDownloadGrid_Click);
            // 
            // NPI_MRP
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1514, 651);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.Panel1);
            this.Name = "NPI_MRP";
            this.Text = "NPI_MRP";
            this.Load += new System.EventHandler(this.NPI_MRP_Load);
            this.Panel1.ResumeLayout(false);
            this.Panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel Panel1;
        private System.Windows.Forms.ComboBox cmbEUShipMeth;
        private System.Windows.Forms.ComboBox cmbEUFOB;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.Button btnDownloadGrid;
        private System.Windows.Forms.TextBox txtProjectName;
        private System.Windows.Forms.Label lblProjectName;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtBuildID;
    }
}