namespace Version3
{
    partial class MRP_Process
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnRunTempMRP = new System.Windows.Forms.Button();
            this.lblTempProcessStatus = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.lblCTFStatus = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.Gainsboro;
            this.groupBox1.Controls.Add(this.panel2);
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Location = new System.Drawing.Point(-3, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(858, 389);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.RoyalBlue;
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Location = new System.Drawing.Point(0, 10);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(858, 38);
            this.panel2.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(425, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(318, 31);
            this.label3.TabIndex = 3;
            this.label3.Text = "|     Process Date          ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(261, 31);
            this.label1.TabIndex = 1;
            this.label1.Text = "Demand Plan Type";
            // 
            // panel1
            // 
            this.panel1.Location = new System.Drawing.Point(0, 44);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(858, 339);
            this.panel1.TabIndex = 3;
            // 
            // btnRunTempMRP
            // 
            this.btnRunTempMRP.BackColor = System.Drawing.Color.Silver;
            this.btnRunTempMRP.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRunTempMRP.Location = new System.Drawing.Point(51, 462);
            this.btnRunTempMRP.Name = "btnRunTempMRP";
            this.btnRunTempMRP.Size = new System.Drawing.Size(210, 76);
            this.btnRunTempMRP.TabIndex = 4;
            this.btnRunTempMRP.Text = "Run Live MRP";
            this.btnRunTempMRP.UseVisualStyleBackColor = false;
            this.btnRunTempMRP.Click += new System.EventHandler(this.btnRunTempMRP_Click);
            // 
            // lblTempProcessStatus
            // 
            this.lblTempProcessStatus.AutoSize = true;
            this.lblTempProcessStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTempProcessStatus.ForeColor = System.Drawing.Color.Peru;
            this.lblTempProcessStatus.Location = new System.Drawing.Point(312, 485);
            this.lblTempProcessStatus.Name = "lblTempProcessStatus";
            this.lblTempProcessStatus.Size = new System.Drawing.Size(405, 29);
            this.lblTempProcessStatus.TabIndex = 13;
            this.lblTempProcessStatus.Text = "Completed  1/1/1900 00:00:00 PM";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.DarkGray;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(51, 576);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(210, 76);
            this.button1.TabIndex = 15;
            this.button1.Text = "Clear to Firm";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblCTFStatus
            // 
            this.lblCTFStatus.AutoSize = true;
            this.lblCTFStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCTFStatus.ForeColor = System.Drawing.Color.Green;
            this.lblCTFStatus.Location = new System.Drawing.Point(312, 599);
            this.lblCTFStatus.Name = "lblCTFStatus";
            this.lblCTFStatus.Size = new System.Drawing.Size(0, 29);
            this.lblCTFStatus.TabIndex = 16;
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.Gray;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(51, 680);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(210, 76);
            this.button2.TabIndex = 17;
            this.button2.Text = "Upload CTF To BQ";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // MRP_Process
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(857, 783);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.lblCTFStatus);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lblTempProcessStatus);
            this.Controls.Add(this.btnRunTempMRP);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MRP_Process";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MRP_Process";
            this.Load += new System.EventHandler(this.MRP_Process_Load);
            this.groupBox1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnRunTempMRP;
        private System.Windows.Forms.Label lblTempProcessStatus;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label lblCTFStatus;
        private System.Windows.Forms.Button button2;
    }
}