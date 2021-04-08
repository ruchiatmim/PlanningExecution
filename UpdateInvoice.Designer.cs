namespace Version3
{
    partial class frmUpdateInvoice
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
            this.btnUpdateBTSData = new Infragistics.Win.Misc.UltraButton();
            this.btnUpdateSBAUS = new Infragistics.Win.Misc.UltraButton();
            this.btnUpdateSBANL = new Infragistics.Win.Misc.UltraButton();
            this.btnUpdateSBASG = new Infragistics.Win.Misc.UltraButton();
            this.SuspendLayout();
            // 
            // btnUpdateBTSData
            // 
            this.btnUpdateBTSData.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpdateBTSData.Location = new System.Drawing.Point(103, 81);
            this.btnUpdateBTSData.Name = "btnUpdateBTSData";
            this.btnUpdateBTSData.Size = new System.Drawing.Size(204, 79);
            this.btnUpdateBTSData.TabIndex = 0;
            this.btnUpdateBTSData.Text = "Update BTS Data";
            this.btnUpdateBTSData.Click += new System.EventHandler(this.ultraButton1_Click);
            // 
            // btnUpdateSBAUS
            // 
            this.btnUpdateSBAUS.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.btnUpdateSBAUS.Location = new System.Drawing.Point(419, 81);
            this.btnUpdateSBAUS.Name = "btnUpdateSBAUS";
            this.btnUpdateSBAUS.Size = new System.Drawing.Size(204, 79);
            this.btnUpdateSBAUS.TabIndex = 1;
            this.btnUpdateSBAUS.Text = "Update SBA - US";
            this.btnUpdateSBAUS.Click += new System.EventHandler(this.btnUpdateSBAUS_Click);
            // 
            // btnUpdateSBANL
            // 
            this.btnUpdateSBANL.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpdateSBANL.Location = new System.Drawing.Point(419, 203);
            this.btnUpdateSBANL.Name = "btnUpdateSBANL";
            this.btnUpdateSBANL.Size = new System.Drawing.Size(204, 79);
            this.btnUpdateSBANL.TabIndex = 2;
            this.btnUpdateSBANL.Text = "Update SBA - NL";
            this.btnUpdateSBANL.Click += new System.EventHandler(this.btnUpdateSBANL_Click);
            // 
            // btnUpdateSBASG
            // 
            this.btnUpdateSBASG.Location = new System.Drawing.Point(419, 330);
            this.btnUpdateSBASG.Name = "btnUpdateSBASG";
            this.btnUpdateSBASG.Size = new System.Drawing.Size(204, 91);
            this.btnUpdateSBASG.TabIndex = 3;
            this.btnUpdateSBASG.Text = "Update SBA - SG";
            this.btnUpdateSBASG.Click += new System.EventHandler(this.btnUpdateSBASG_Click);
            // 
            // frmUpdateInvoice
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(929, 611);
            this.Controls.Add(this.btnUpdateSBASG);
            this.Controls.Add(this.btnUpdateSBANL);
            this.Controls.Add(this.btnUpdateSBAUS);
            this.Controls.Add(this.btnUpdateBTSData);
            this.Name = "frmUpdateInvoice";
            this.Text = "Update Invoice";
            this.ResumeLayout(false);

        }

        #endregion

        private Infragistics.Win.Misc.UltraButton btnUpdateBTSData;
        private Infragistics.Win.Misc.UltraButton btnUpdateSBAUS;
        private Infragistics.Win.Misc.UltraButton btnUpdateSBANL;
        private Infragistics.Win.Misc.UltraButton btnUpdateSBASG;
    }
}