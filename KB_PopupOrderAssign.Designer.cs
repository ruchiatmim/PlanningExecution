namespace Version3
{
    partial class KB_PopupOrderAssign
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
            this.lblPartNum = new System.Windows.Forms.Label();
            this.lblHeading = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.pnlOrder = new System.Windows.Forms.Panel();
            this.Label1 = new System.Windows.Forms.Label();
            this.btnDone = new System.Windows.Forms.Button();
            this.Panel1.SuspendLayout();
            this.pnlOrder.SuspendLayout();
            this.SuspendLayout();
            // 
            // Panel1
            // 
            this.Panel1.BackColor = System.Drawing.Color.DarkGray;
            this.Panel1.Controls.Add(this.lblPartNum);
            this.Panel1.Controls.Add(this.lblHeading);
            this.Panel1.Controls.Add(this.label4);
            this.Panel1.Controls.Add(this.Label2);
            this.Panel1.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Panel1.Location = new System.Drawing.Point(3, 2);
            this.Panel1.Name = "Panel1";
            this.Panel1.Size = new System.Drawing.Size(358, 28);
            this.Panel1.TabIndex = 10;
            // 
            // lblPartNum
            // 
            this.lblPartNum.AutoSize = true;
            this.lblPartNum.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPartNum.ForeColor = System.Drawing.Color.Black;
            this.lblPartNum.Location = new System.Drawing.Point(253, 4);
            this.lblPartNum.Name = "lblPartNum";
            this.lblPartNum.Size = new System.Drawing.Size(0, 16);
            this.lblPartNum.TabIndex = 7;
            this.lblPartNum.UseWaitCursor = true;
            // 
            // lblHeading
            // 
            this.lblHeading.AutoSize = true;
            this.lblHeading.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHeading.ForeColor = System.Drawing.Color.Black;
            this.lblHeading.Location = new System.Drawing.Point(97, 4);
            this.lblHeading.Name = "lblHeading";
            this.lblHeading.Size = new System.Drawing.Size(0, 16);
            this.lblHeading.TabIndex = 7;
            this.lblHeading.UseWaitCursor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(184, 4);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(69, 16);
            this.label4.TabIndex = 5;
            this.label4.Text = "Part Num :";
            this.label4.UseWaitCursor = true;
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label2.ForeColor = System.Drawing.Color.Black;
            this.Label2.Location = new System.Drawing.Point(9, 4);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(85, 16);
            this.Label2.TabIndex = 5;
            this.Label2.Text = "KB Location :";
            this.Label2.UseWaitCursor = true;
            // 
            // pnlOrder
            // 
            this.pnlOrder.AutoScroll = true;
            this.pnlOrder.Controls.Add(this.Label1);
            this.pnlOrder.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pnlOrder.Location = new System.Drawing.Point(3, 31);
            this.pnlOrder.Name = "pnlOrder";
            this.pnlOrder.Size = new System.Drawing.Size(358, 223);
            this.pnlOrder.TabIndex = 11;
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label1.ForeColor = System.Drawing.Color.Black;
            this.Label1.Location = new System.Drawing.Point(38, 3);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(108, 16);
            this.Label1.TabIndex = 9;
            this.Label1.Text = "Order Number";
            // 
            // btnDone
            // 
            this.btnDone.BackColor = System.Drawing.Color.Green;
            this.btnDone.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDone.ForeColor = System.Drawing.Color.White;
            this.btnDone.Location = new System.Drawing.Point(279, 260);
            this.btnDone.Name = "btnDone";
            this.btnDone.Size = new System.Drawing.Size(82, 24);
            this.btnDone.TabIndex = 13;
            this.btnDone.Text = "Done";
            this.btnDone.UseVisualStyleBackColor = false;
            this.btnDone.Click += new System.EventHandler(this.btnDone_Click);
            // 
            // KB_PopupOrderAssign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(363, 286);
            this.Controls.Add(this.btnDone);
            this.Controls.Add(this.pnlOrder);
            this.Controls.Add(this.Panel1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "KB_PopupOrderAssign";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "BIN ORDER";
            this.Panel1.ResumeLayout(false);
            this.Panel1.PerformLayout();
            this.pnlOrder.ResumeLayout(false);
            this.pnlOrder.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Panel Panel1;
        internal System.Windows.Forms.Label lblHeading;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.Panel pnlOrder;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Label lblPartNum;
        internal System.Windows.Forms.Label label4;
        internal System.Windows.Forms.Button btnDone;
    }
}