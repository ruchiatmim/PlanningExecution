namespace Version3
{
    partial class KanbanOrderAssign
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
            this.Label20 = new System.Windows.Forms.Label();
            this.PnlPartNum = new System.Windows.Forms.Panel();
            this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.kanbanLayoutSecition1 = new KanbanLayoutSecition();
            this.SuspendLayout();
            // 
            // Label20
            // 
            this.Label20.AutoSize = true;
            this.Label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label20.ForeColor = System.Drawing.Color.Black;
            this.Label20.Location = new System.Drawing.Point(1662, 9);
            this.Label20.Name = "Label20";
            this.Label20.Size = new System.Drawing.Size(99, 20);
            this.Label20.TabIndex = 107;
            this.Label20.Text = "PART NUM";
            // 
            // PnlPartNum
            // 
            this.PnlPartNum.AutoScroll = true;
            this.PnlPartNum.BackColor = System.Drawing.Color.WhiteSmoke;
            this.PnlPartNum.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PnlPartNum.ForeColor = System.Drawing.Color.Black;
            this.PnlPartNum.Location = new System.Drawing.Point(1665, 31);
            this.PnlPartNum.Name = "PnlPartNum";
            this.PnlPartNum.Size = new System.Drawing.Size(156, 764);
            this.PnlPartNum.TabIndex = 106;
            // 
            // kanbanLayoutSecition1
            // 
            this.kanbanLayoutSecition1.BackColor = System.Drawing.Color.White;
            this.kanbanLayoutSecition1.Location = new System.Drawing.Point(0, 0);
            this.kanbanLayoutSecition1.Name = "kanbanLayoutSecition1";
            this.kanbanLayoutSecition1.Size = new System.Drawing.Size(1564, 943);
            this.kanbanLayoutSecition1.TabIndex = 0;
            // 
            // KanbanOrderAssign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1972, 927);
            this.Controls.Add(this.Label20);
            this.Controls.Add(this.PnlPartNum);
            this.Controls.Add(this.kanbanLayoutSecition1);
            this.Name = "KanbanOrderAssign";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Kanban Order Assign";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private KanbanLayoutSecition kanbanLayoutSecition1;
        internal System.Windows.Forms.Label Label20;
        internal System.Windows.Forms.Panel PnlPartNum;
        private System.Windows.Forms.ToolTip ToolTip1;
    }
}