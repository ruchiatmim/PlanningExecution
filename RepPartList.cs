using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;

namespace Version3
{
    public partial class RepPartList : Form
    {
        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        public RepPartList()
        {
            InitializeComponent();
            loadData();
         
        }


        public void loadData()
        {
            string qry = "select MRPPartNum, PlanGroup, AsmCategory,PlanningItem,AttachRate  from Waterfall_Rep_PartsList  order by MRPPartNum";
            DataSet ds = getPartList(qry);       
            ultraGridRepPartList.DataSource = ds.Tables[0];           
            ultraGridRepPartList.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
            lblNumRecKan.Text = "Number of Records : " + ds.Tables[0].Rows.Count.ToString();
            ultraGridRepPartList.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            //this.ultraGridRepPartList.DisplayLayout.Bands[0].Columns["Delete"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.EditButton;
            //this.ultraGridRepPartList.Rows.Band.Columns["Delete"].ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
            //Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Delete");
            //ultraGridRepPartList.DisplayLayout.Bands[0].Columns.Add(ultraGridColumn51);

            Infragistics.Win.UltraWinGrid.UltraGridColumn c1 = ultraGridRepPartList.DisplayLayout.Bands[0].Columns["MRPPartNum"];
           // c1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

            Infragistics.Win.UltraWinGrid.UltraGridColumn c2 = ultraGridRepPartList.DisplayLayout.Bands[0].Columns["PlanGroup"];
          //  c2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

            Infragistics.Win.UltraWinGrid.UltraGridColumn c3 = ultraGridRepPartList.DisplayLayout.Bands[0].Columns["AsmCategory"];
            //c3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;            
        }


        public DataSet getPartList(string sqlQry)
        {       
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlQry, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;

        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            this.ultraGridRepPartList.DisplayLayout.Bands[0].AddNew();
            this.ultraGridRepPartList.Rows[ultraGridRepPartList.Rows.Count - 1].Selected = true;
           
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            if (ultraGridRepPartList.Selected.Rows.Count > 0)
            {
                string PartNum = ultraGridRepPartList.Rows[ultraGridRepPartList.Selected.Rows[0].Index].Cells["MRPPartNum"].Value.ToString();
                string qry = "delete  from Waterfall_Rep_PartsList where MRPPartNum='" + PartNum + "' select @@error";
                DataSet ds = getPartList(qry);
                if (ds.Tables[0].Rows[0][0].ToString()=="0")
                    MessageBox.Show("Part has been deleted successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);                
                else
                    MessageBox.Show("Not Saved. Please check the error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                loadData();         
            }
        }

        private void releasesBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            string saveQuery = " ";
            string strMRPPartNum  ="";
            string strPlanGroup ="";
            string strAsmCategory ="";

            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGridRepPartList.Rows)
            {                
                strMRPPartNum  =dr.Cells["MRPPArtNum"].Value.ToString();
                strPlanGroup = dr.Cells["PlanGroup"].Value.ToString();
                 strAsmCategory = dr.Cells["AsmCategory"].Value.ToString();
                 if (strMRPPartNum != "" && strPlanGroup != "" && strAsmCategory != "")
                 {
                     saveQuery = saveQuery + "   exec   sp_saveRepPartList '" + strMRPPartNum + "','" + strPlanGroup + "','" + strAsmCategory + "'";
                 }
            }
            DataSet ds = getPartList(saveQuery);
            SqlConnection cn = new SqlConnection(constr);
            
            cn.Open();
            SqlTransaction tran = cn.BeginTransaction("Transaction");
            try
            {
                SqlCommand cmd = new SqlCommand(saveQuery, cn);              
                cmd.Transaction = tran;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                cmd.ExecuteNonQuery();
                tran.Commit();
                MessageBox.Show("Part has been saved successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exp)
            {
                tran.Rollback();
                MessageBox.Show("Not Saved. Please check the error" + exp.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
            cn.Close();
            loadData();         
        }

        private void ultraGridRepPartList_ClickCell(object sender, Infragistics.Win.UltraWinGrid.ClickCellEventArgs e)
        {

        }

        private void btnRefreshGrid_Click(object sender, EventArgs e)
        {
            loadData();
        }

        private void ultraGridRepPartList_InitializeTemplateAddRow(object sender, Infragistics.Win.UltraWinGrid.InitializeTemplateAddRowEventArgs e)
        {
          
        }

        private void ultraGridRepPartList_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
            //Infragistics.Win.UltraWinGrid.UltraGridLayout layout = e.Layout;
            //Infragistics.Win.UltraWinGrid.UltraGridBand band = layout.Bands[0];
            //Infragistics.Win.UltraWinGrid.UltraGridColumn imageColumn = band.Columns.Add("Del");
            //imageColumn.DataType = typeof(Image);            
            //imageColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Image;           
        }

        private void ultraGridRepPartList_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
        {
            //if (e.Row.IsDataRow)
            //{
            //    // Create an image from the path string in the "Path" cell
            //    string path="E:\\Ruchi\\NewPlanningExceution\\MaintainGoogleDoc\\MaintainGoogleDoc\\Resources\\delete2.ico";
            //    Image image = Bitmap.FromFile(path);

            //    // Put the image in the "Image" cell
            //    e.Row.Cells["Del"].Value = image;
            //}
        }

        private void btnKanAddNew_Click(object sender, EventArgs e)
        {
            clear();
        }

        private void clear()
        {
            textAsmCategory.Text = "";
            textPartNum.Text = "";
            textPlanGroup.Text = "";
            textboxPlanningItem.Text = "";
            textAttachRate.Value = 0.00;
        }

        private void btnKanSave_Click(object sender, EventArgs e)
        {

            string saveQuery = " ";
            string strMRPPartNum = "";
            string strPlanGroup = "";
            string strAsmCategory = "";
            string strPlanningItem = "";
            decimal decAttachRate;

            strMRPPartNum = textPartNum.Text.Trim();
            strPlanGroup = textPlanGroup.Text.ToString();
            strAsmCategory = textAsmCategory.Text.Trim();
            strPlanningItem = textboxPlanningItem.Text.Trim();
            decAttachRate = Convert.ToDecimal(textAttachRate.Value);
                if (strMRPPartNum != "" && strPlanGroup != "" && strAsmCategory != "")
                {
                    saveQuery = "   exec   sp_saveRepPartList '" + strMRPPartNum + "','" + strPlanGroup + "','" + strAsmCategory + "','" + strPlanningItem + "','" + decAttachRate.ToString() + "'";
                }
          
            DataSet ds = getPartList(saveQuery);

            SqlConnection cn = new SqlConnection(constr);

          cn.Open();
          //  SqlTransaction tran = cn.BeginTransaction("Transaction");
            try
            {
                SqlCommand cmd = new SqlCommand(saveQuery, cn);

             //   cmd.Transaction = tran;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                cmd.ExecuteNonQuery();
              //  tran.Commit();
                MessageBox.Show("Part has been saved successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exp)
            {
               // tran.Rollback();
                MessageBox.Show("Not Saved. Please check the error" + exp.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            cn.Close();
            clear();
            loadData();         
            
        }

        private void ultraGridRepPartList_AfterSelectChange(object sender, Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs e)
        {
            if (ultraGridRepPartList.Selected.Rows.Count > 0)
            {
                textPartNum.Text = ultraGridRepPartList.Selected.Rows[0].Cells["MRPPartNum"].Value.ToString();
               
                this.textPlanGroup.Text= ultraGridRepPartList.Selected.Rows[0].Cells["PlanGroup"].Value.ToString();
                this.textAsmCategory.Text = ultraGridRepPartList.Selected.Rows[0].Cells["AsmCategory"].Value.ToString();
                this.textboxPlanningItem.Text = ultraGridRepPartList.Selected.Rows[0].Cells["PlanningItem"].Value.ToString();
                if (ultraGridRepPartList.Selected.Rows[0].Cells["AttachRate"].Value.ToString() != "")
                    this.textAttachRate.Value = Convert.ToDecimal(ultraGridRepPartList.Selected.Rows[0].Cells["AttachRate"].Value.ToString());
                else
                    this.textAttachRate.Value = 0.00;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string qry = "delete  from Waterfall_Rep_PartsList where MRPPartNum='" +textPartNum.Text.Trim() + "' select @@error";
            DataSet ds = getPartList(qry);
            if (ds.Tables[0].Rows[0][0].ToString() == "0")
                MessageBox.Show("Part has been deleted successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show("Not Saved. Please check the error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            loadData();
            clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textAsmCategory.Text = textAsmCategory.Text + "|";
            textAsmCategory.Focus();
            textAsmCategory.SelectionStart = textAsmCategory.Text.Length; // add some logic if length is 0
            textAsmCategory.SelectionLength = 0;         
        }

       
    }
}
