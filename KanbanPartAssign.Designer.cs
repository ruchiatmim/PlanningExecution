namespace Version3
{
    partial class KanbanPartAssign
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
            this.components = new System.ComponentModel.Container();
            this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.PnlPartNum = new System.Windows.Forms.Panel();
            this.Label20 = new System.Windows.Forms.Label();
            this.kanbanLayoutSecition1 = new KanbanLayoutSecition();
            this.SuspendLayout();
            // 
            // PnlPartNum
            // 
            this.PnlPartNum.AutoScroll = true;
            this.PnlPartNum.BackColor = System.Drawing.Color.WhiteSmoke;
            this.PnlPartNum.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PnlPartNum.ForeColor = System.Drawing.Color.Black;
            this.PnlPartNum.Location = new System.Drawing.Point(1687, 31);
            this.PnlPartNum.Name = "PnlPartNum";
            this.PnlPartNum.Size = new System.Drawing.Size(148, 764);
            this.PnlPartNum.TabIndex = 104;
            this.PnlPartNum.Paint += new System.Windows.Forms.PaintEventHandler(this.rt_Paint);
            // 
            // Label20
            // 
            this.Label20.AutoSize = true;
            this.Label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label20.ForeColor = System.Drawing.Color.Black;
            this.Label20.Location = new System.Drawing.Point(1683, 9);
            this.Label20.Name = "Label20";
            this.Label20.Size = new System.Drawing.Size(99, 20);
            this.Label20.TabIndex = 105;
            this.Label20.Text = "PART NUM";
            // 
            // kanbanLayoutSecition1
            // 
            this.kanbanLayoutSecition1.BackColor = System.Drawing.Color.White;
            this.kanbanLayoutSecition1.Location = new System.Drawing.Point(1, 2);
            this.kanbanLayoutSecition1.Name = "kanbanLayoutSecition1";
            this.kanbanLayoutSecition1.Size = new System.Drawing.Size(1567, 943);
            this.kanbanLayoutSecition1.TabIndex = 106;
            // 
            // KanbanPartAssign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1916, 1040);
            this.Controls.Add(this.kanbanLayoutSecition1);
            this.Controls.Add(this.Label20);
            this.Controls.Add(this.PnlPartNum);
            this.Name = "KanbanPartAssign";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "KanbanPartAssign";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.ToolTip ToolTip1;
        internal System.Windows.Forms.Panel PnlPartNum;
        internal System.Windows.Forms.Label Label20;
        private KanbanLayoutSecition kanbanLayoutSecition1;


    }
}