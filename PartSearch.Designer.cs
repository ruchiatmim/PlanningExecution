namespace Version3
{
    partial class PartSearch
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            this.Panel1 = new System.Windows.Forms.Panel();
            this.txtManfPartNum = new System.Windows.Forms.TextBox();
            this.lblFilter = new System.Windows.Forms.Label();
            this.btnSearch = new System.Windows.Forms.Button();
            this.GroupBox2 = new System.Windows.Forms.GroupBox();
            this.lblDescription = new System.Windows.Forms.Label();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.lblScannedPart = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.gridview = new System.Windows.Forms.DataGridView();
            this.Panel1.SuspendLayout();
            this.GroupBox2.SuspendLayout();
            this.GroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridview)).BeginInit();
            this.SuspendLayout();
            // 
            // Panel1
            // 
            this.Panel1.AutoSize = true;
            this.Panel1.BackColor = System.Drawing.SystemColors.Desktop;
            this.Panel1.Controls.Add(this.txtManfPartNum);
            this.Panel1.Controls.Add(this.lblFilter);
            this.Panel1.Controls.Add(this.btnSearch);
            this.Panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.Panel1.Location = new System.Drawing.Point(0, 0);
            this.Panel1.Name = "Panel1";
            this.Panel1.Size = new System.Drawing.Size(720, 53);
            this.Panel1.TabIndex = 2;
            // 
            // txtManfPartNum
            // 
            this.txtManfPartNum.Location = new System.Drawing.Point(208, 24);
            this.txtManfPartNum.Name = "txtManfPartNum";
            this.txtManfPartNum.Size = new System.Drawing.Size(224, 20);
            this.txtManfPartNum.TabIndex = 5;
    //        this.txtManfPartNum.TextChanged += new System.EventHandler(this.txtManfPartNum_TextChanged);
            // 
            // lblFilter
            // 
            this.lblFilter.AutoSize = true;
            this.lblFilter.BackColor = System.Drawing.SystemColors.Desktop;
            this.lblFilter.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFilter.ForeColor = System.Drawing.Color.White;
            this.lblFilter.Location = new System.Drawing.Point(12, 24);
            this.lblFilter.Name = "lblFilter";
            this.lblFilter.Size = new System.Drawing.Size(190, 18);
            this.lblFilter.TabIndex = 4;
            this.lblFilter.Text = "Manufacture Part Number";
            // 
            // btnSearch
            // 
            this.btnSearch.BackColor = System.Drawing.Color.DarkRed;
            this.btnSearch.ForeColor = System.Drawing.Color.White;
            this.btnSearch.Location = new System.Drawing.Point(455, 16);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(90, 34);
            this.btnSearch.TabIndex = 1;
            this.btnSearch.Text = "Search";
            this.btnSearch.UseVisualStyleBackColor = false;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // GroupBox2
            // 
            this.GroupBox2.AutoSize = true;
            this.GroupBox2.BackColor = System.Drawing.SystemColors.Control;
            this.GroupBox2.Controls.Add(this.lblDescription);
            this.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.GroupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GroupBox2.ForeColor = System.Drawing.Color.Black;
            this.GroupBox2.Location = new System.Drawing.Point(0, 275);
            this.GroupBox2.Name = "GroupBox2";
            this.GroupBox2.Size = new System.Drawing.Size(720, 208);
            this.GroupBox2.TabIndex = 9;
            this.GroupBox2.TabStop = false;
            this.GroupBox2.Text = "Part Description";
            // 
            // lblDescription
            // 
            this.lblDescription.AutoSize = true;
            this.lblDescription.Location = new System.Drawing.Point(9, 31);
            this.lblDescription.Name = "lblDescription";
            this.lblDescription.Size = new System.Drawing.Size(0, 13);
            this.lblDescription.TabIndex = 0;
            // 
            // GroupBox1
            // 
            this.GroupBox1.AutoSize = true;
            this.GroupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.GroupBox1.Controls.Add(this.lblScannedPart);
            this.GroupBox1.Controls.Add(this.Label1);
            this.GroupBox1.Controls.Add(this.gridview);
            this.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.GroupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GroupBox1.ForeColor = System.Drawing.Color.Black;
            this.GroupBox1.Location = new System.Drawing.Point(0, 53);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(720, 222);
            this.GroupBox1.TabIndex = 8;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Text = "Part Information";
            // 
            // lblScannedPart
            // 
            this.lblScannedPart.AutoSize = true;
            this.lblScannedPart.Location = new System.Drawing.Point(125, 30);
            this.lblScannedPart.Name = "lblScannedPart";
            this.lblScannedPart.Size = new System.Drawing.Size(0, 13);
            this.lblScannedPart.TabIndex = 2;
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(9, 30);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(88, 13);
            this.Label1.TabIndex = 1;
            this.Label1.Text = "Scanned Part:";
            // 
            // gridview
            // 
            this.gridview.AllowUserToResizeColumns = false;
            this.gridview.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.SaddleBrown;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.SaddleBrown;
            this.gridview.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.gridview.BackgroundColor = System.Drawing.SystemColors.Control;
            this.gridview.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.gridview.CausesValidation = false;
            this.gridview.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.gridview.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gridview.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.gridview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.SaddleBrown;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.gridview.DefaultCellStyle = dataGridViewCellStyle3;
            this.gridview.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.gridview.EnableHeadersVisualStyles = false;
            this.gridview.GridColor = System.Drawing.SystemColors.Control;
            this.gridview.Location = new System.Drawing.Point(12, 58);
            this.gridview.MultiSelect = false;
            this.gridview.Name = "gridview";
            this.gridview.ReadOnly = true;
            this.gridview.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gridview.RowHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.gridview.RowHeadersVisible = false;
            dataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle5.ForeColor = System.Drawing.Color.SaddleBrown;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.Color.SaddleBrown;
            this.gridview.RowsDefaultCellStyle = dataGridViewCellStyle5;
            this.gridview.RowTemplate.DefaultCellStyle.BackColor = System.Drawing.SystemColors.Control;
            this.gridview.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.SaddleBrown;
            this.gridview.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Control;
            this.gridview.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.SaddleBrown;
            this.gridview.RowTemplate.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gridview.RowTemplate.ReadOnly = true;
            this.gridview.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.gridview.ShowCellErrors = false;
            this.gridview.ShowCellToolTips = false;
            this.gridview.ShowEditingIcon = false;
            this.gridview.ShowRowErrors = false;
            this.gridview.Size = new System.Drawing.Size(745, 145);
            this.gridview.TabIndex = 0;
            this.gridview.TabStop = false;
            this.gridview.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gridview_CellMouseClick);
            this.gridview.MouseClick += new System.Windows.Forms.MouseEventHandler(this.gridview_MouseClick);
            this.gridview.CellMouseDown += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gridview_CellMouseDown);
            this.gridview.CellStateChanged += new System.Windows.Forms.DataGridViewCellStateChangedEventHandler(this.gridview_CellStateChanged);
            this.gridview.Click += new System.EventHandler(this.gridview_Click);
          //  this.gridview.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridview_CellContentClick);
            // 
            // PartSearch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(720, 483);
            this.Controls.Add(this.GroupBox2);
            this.Controls.Add(this.GroupBox1);
            this.Controls.Add(this.Panel1);
            this.Name = "PartSearch";
            this.Text = "PartSearch";
            this.Panel1.ResumeLayout(false);
            this.Panel1.PerformLayout();
            this.GroupBox2.ResumeLayout(false);
            this.GroupBox2.PerformLayout();
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridview)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.Panel Panel1;
        internal System.Windows.Forms.TextBox txtManfPartNum;
        internal System.Windows.Forms.Label lblFilter;
        internal System.Windows.Forms.Button btnSearch;
        internal System.Windows.Forms.GroupBox GroupBox2;
        internal System.Windows.Forms.Label lblDescription;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.Label lblScannedPart;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.DataGridView gridview;
    }
}