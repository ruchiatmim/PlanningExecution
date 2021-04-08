using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using Infragistics.Win.UltraWinEditors;
using Infragistics.Win.UltraWinGrid;

namespace Version3
{
    public partial class RegionOption : Form
    {
        public RegionOption()
        {
            InitializeComponent();
            showRegionOption();
        }

        private void showRegionOption()
        {
            grdRegionOption.DataSource = null;

            try
            {                             
                string sqlGrid = " SELECT [OptionPN], [ASIA], [EUROPE], [NORTH AMERICA], [SOUTH AMERICA], [TAIWAN]  FROM [MIMDIST].[dbo].[MRP_C_RegionOption] "; 
                grdRegionOption.DataSource = getDataSet(sqlGrid).Tables[0];             

            }
            catch
            {
       
            }
           
        }

        public DataSet getDataSet(string sqlStr)
        {        
           
            string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlStr, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;

        }

        private void grdRegionOption_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {         
            UltraGridLayout layout = e.Layout;
            UltraGridBand band = layout.Bands[0];
            UltraGridColumn imageColumn = band.Columns.Add("Delete");
            imageColumn.DataType = typeof(Image);
            imageColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Image;

        }

        private void grdRegionOption_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            if (e.Row.Cells.Exists("Delete"))
            {                
                string path = "image\\delete2.ico";
                Image image = Bitmap.FromFile(path);
                e.Row.Cells["Delete"].Value = image;
            }
        }


        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            this.grdRegionOption.DisplayLayout.Bands[0].AddNew();
            this.grdRegionOption.Rows[grdRegionOption.Rows.Count - 1].Selected = true;
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
           
        }

        private void grdRegionOption_ClickCell(object sender, ClickCellEventArgs e)
        {
            if (e.Cell.Column.Index == 6)
            {
                if (e.Cell.Row.Cells["OptionPN"].Value.ToString() != "")
                {
                    DeleteRowDataBase(e.Cell.Row.Cells["OptionPN"].Value.ToString());
                }

                this.grdRegionOption.Rows[e.Cell.Row.Index].Delete();
            }
        }

        public void DeleteRowDataBase(string OptionPN)
        {            
            
            string sqlDelete = "delete from [MIMDIST].[dbo].[MRP_C_RegionOption]  where OptionPN=" + OptionPN.ToString();
            getDataSet(sqlDelete);

        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            string sql = "";
            bool duplicatePart = false;
              bool emptyPart = false;
            string OptionPN_list = "";
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.grdRegionOption.Rows)
            {
                string strOptionPN = "";
                if (row.Cells["OptionPN"].Text.ToString().Trim() != "")
                    strOptionPN = row.Cells["OptionPN"].Text.ToString();
                string strASIA = row.Cells["ASIA"].Text.ToString();
                string strEUROPE = row.Cells["EUROPE"].Text.ToString();
                string strNORTH_AMERICA = row.Cells["NORTH AMERICA"].Text.ToString();
                string strSOUTH_AMERICA = row.Cells["SOUTH AMERICA"].Text.ToString();
                string strTAIWAN = row.Cells["TAIWAN"].Text.ToString();              
             
                if (OptionPN_list.Contains(strOptionPN))
                    duplicatePart = true;

                OptionPN_list = OptionPN_list + strOptionPN + ",";
                if (!duplicatePart)
                {
                    if (strOptionPN != "" && strASIA != "" && strEUROPE != "" && strNORTH_AMERICA != "" && strSOUTH_AMERICA != "" && strTAIWAN != "")
                        sql = sql + "  insert into [dbo].[MRP_C_RegionOption](OptionPN,ASIA,EUROPE,[NORTH AMERICA],[SOUTH AMERICA],TAIWAN) values ('" + strOptionPN + "','" + strASIA + "','" + strEUROPE + "','" + strNORTH_AMERICA + "','" + strSOUTH_AMERICA + "','" + strTAIWAN + "' )  ";
                    else
                        emptyPart = true;
                }               
            }

            if (duplicatePart)
                MessageBox.Show("Please check for duplicate part");
            else if (emptyPart)
                MessageBox.Show("Please check for empty part");
            else
            {
                if (sql != "")
                {
                    sql = " Delete from [dbo].[MRP_C_RegionOption]   " + sql;
                    string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
                    SqlConnection cn = new SqlConnection(constr);
                    cn.Open();
                    SqlTransaction tran = cn.BeginTransaction("Transaction");
                    SqlCommand cmd = new SqlCommand(sql, cn);
                    cmd.Transaction = tran;
                    try
                    {
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataSet ds = new DataSet();
                        da.Fill(ds);
                        tran.Commit();
                        MessageBox.Show("Updated successfully");
                    }
                    catch (Exception exp)
                    {
                        string error = exp.Message.ToString();
                        tran.Rollback();
                        MessageBox.Show("Failed to Update. Please check the data");
                    }
                    cn.Close();
                }
            }
                showRegionOption();
            }

        private void grdRegionOption_Leave(object sender, EventArgs e)
        {

        }

        private void grdRegionOption_CellChange(object sender, CellEventArgs e)
        {
             string optionPartNum="";
             if ((e.Cell.Column.Index == 0) && (e.Cell.Row.Cells["OptionPN"].Value != null))
             {
                 optionPartNum = e.Cell.Row.Cells["OptionPN"].Value.ToString();
             }             

        }
       
    }
}
