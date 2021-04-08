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
using System.IO;

using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Diagnostics;

namespace Version3
{
    public partial class VendorTrackingNum : Form
    {

          string constrMVERP = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
          string constrEpicor = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;


        public VendorTrackingNum()
        {
     
            InitializeComponent();         
            try
            {
                loadVendor();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }
        
            try
            {
                loadCarier();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }
        
        }

        private void btnFilter_Click(object sender, EventArgs e)
        {
            bindgrid();
        }

        public void clearALL()
        {
            dtRelease = null;
            this.CmbCarrier.Text = "";
            txtPartNum.Text = "";
            txtPO.Text = "";
            this.cmbVendor.Text = "";
           
            this.vendorGrid.DataSource = null;
        }
        
        public void loadVendor()
        {
            string strqry = " SELECT distinct vendor_code FROM [MIMDIST].[dbo].[apmaster] where status_type=5 and [vend_class_code]='VENDOR'      SELECT distinct part_no FROM [MIMDIST].[dbo].[inv_master] WHERE     (void = 'N') AND (buyer = 'SOI') AND (obsolete = 0)";
            //string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
            DataSet ds = getDataSet(strqry, constrMVERP);
            DataTable dtVendor= ds.Tables[0];
            DataRow dr = dtVendor.NewRow();          
            dr[0] = "Select";
            dtVendor.Rows.Add(dr);
            this.cmbVendor.DataSource = dtVendor;
            cmbVendor.ValueMember = "vendor_code";
            cmbVendor.DisplayMember = "vendor_code";     
            
            
            DataTable dtLoc=new DataTable();
          //  DataColumn col = new DataColumn("Location");
            dtLoc.Columns.Add(new DataColumn("Location"));
            DataRow drLoc = dtLoc.NewRow();
            drLoc[0] = "ALL";
            dtLoc.Rows.Add(drLoc);

            drLoc = dtLoc.NewRow();
            drLoc[0] = "AT";
            dtLoc.Rows.Add(drLoc);

            drLoc = dtLoc.NewRow();
            drLoc[0] = "NL";
            dtLoc.Rows.Add(drLoc);

            drLoc = dtLoc.NewRow();
            drLoc[0] = "SG";
            dtLoc.Rows.Add(drLoc);

           

       //     this.cmbLocation.DataSource = dtLoc;
        //    cmbLocation.ValueMember = "Location";
       //     cmbLocation.DisplayMember = "Location";
       //     cmbLocation.SelectedText = "ALL";
            
            
        }

        public void loadCarier()
        {
            
          //  string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
          DataTable  dtCarrier = getDataSet("SELECT ship_via_code, ship_via_name FROM arshipv", constrMVERP).Tables[0];
          DataRow dr = dtCarrier.NewRow();
                dr[0]=0;
            dr[1]="Select";
           // dtCarrier.Rows.Add(dr);
            if (CmbCarrier.DataSource == null)
            {
                CmbCarrier.DisplayMember = "ship_via_name";
                CmbCarrier.ValueMember = "ship_via_code";
                CmbCarrier.DataSource = dtCarrier;
            }
            CmbCarrier.Rows.Band.Columns["ship_via_code"].Hidden = true;
        }

        DataTable dtRelease; 
        public void bindgrid()
        {
            //clearALL();
           // string sql = " SELECT  dbo.purchase.vendor_no AS [Vendor No], dbo.releases.po_no AS [PO #], dbo.releases.po_line AS [PO Line], dbo.releases.row_id AS [Release #],  dbo.releases.part_no AS [Part Num], cast(dbo.releases.quantity as integer) AS Quantity, cast((dbo.releases.quantity - dbo.releases.received) as integer) AS [Due Quantity], dbo.arshipv.ship_via_name AS Carrier, dbo.releases.m_tracking AS Tracking FROM         dbo.releases INNER JOIN                       dbo.pur_list ON dbo.releases.po_no = dbo.pur_list.po_no AND dbo.releases.po_line = dbo.pur_list.line INNER JOIN   dbo.purchase ON dbo.pur_list.po_no = dbo.purchase.po_no INNER JOIN  dbo.arshipv ON ISNULL(dbo.releases.carrier, dbo.purchase.ship_via) = dbo.arshipv.ship_via_code  WHERE     (dbo.purchase.status = 'O') AND (dbo.purchase.m_po_type = 3) AND (dbo.purchase.void = 'N') AND (dbo.purchase.date_of_order > GETDATE() - 180) ";
            //if  ( this.txtVendor.Text.Trim() != "")
            //    sql = sql + " and vendor_no='" + this.txtVendor.Text.Trim() + "'";



            string sql = "SELECT     dbo.purchase.vendor_no AS [Vendor No], dbo.releases.po_no AS [PO #], dbo.releases.po_line AS [PO Line], dbo.releases.row_id AS [Release #],                       dbo.releases.part_no AS [Part Num], CAST(dbo.releases.quantity AS integer) AS Quantity, CAST(dbo.releases.quantity - dbo.releases.received AS integer)                       AS [Due Quantity], dbo.arshipv.ship_via_name AS Carrier, dbo.releases_tracking.Tracking, substring(dbo.releases.location,1,2) as Loc FROM         dbo.releases INNER JOIN                      dbo.pur_list ON dbo.releases.po_no = dbo.pur_list.po_no AND dbo.releases.po_line = dbo.pur_list.line INNER JOIN                       dbo.purchase ON dbo.pur_list.po_no = dbo.purchase.po_no INNER JOIN                       dbo.arshipv ON ISNULL(dbo.releases.carrier, dbo.purchase.ship_via) = dbo.arshipv.ship_via_code LEFT OUTER JOIN                       dbo.releases_tracking ON dbo.releases.row_id = dbo.releases_tracking.row_id WHERE     (dbo.purchase.status = 'O') AND (dbo.purchase.m_po_type = 3) AND (dbo.purchase.void = 'N') AND (dbo.purchase.date_of_order > GETDATE() - 180)";
            
            this.vendorGrid.DataSource = null;

         //  if (this.cmbLocation.Text.Trim() != "ALL")
            //    sql = sql + " and substring(dbo.releases.location,1,2)='" + cmbLocation.Text.ToString() + "'";


                if (this.txtPO.Text.Trim() != "")
                    sql = sql + " and dbo.releases.po_no='" + txtPO.Text.Trim() + "'";
                if (this.txtPartNum.Text.Trim() != "")
                    sql = sql + " and  dbo.releases.part_no like '" + txtPartNum.Text.Trim() + "%'";
                if (this.cmbVendor.Text.Trim() != "")
                    sql = sql + " and dbo.purchase.vendor_no='" + cmbVendor.Text.Trim() + "'";
                if (this.CmbCarrier.Text.Trim() != "")
                   sql = sql + " and ISNULL(dbo.releases.carrier, dbo.purchase.ship_via)='" + CmbCarrier.Value.ToString() + "'";


            
                sql = sql + " ORDER BY dbo.releases.m_factory_date Asc ";
               
                DataSet dsRelease = getDataSet(sql, this.constrEpicor);
                dtRelease = dsRelease.Tables[0];//.Select("po_line=1").CopyToDataTable();
                bindingSourceRelease.DataSource = dtRelease;
            
                //comboBox1.DataBindings.Add("SelectedItem", releasesBindingSource, "OverflowBehaviour", false, DataSourceUpdateMode.OnValidation);
                if (dtRelease != null)
                {
                    if (dtRelease.Rows.Count > 0)
                    {
                        try
                        {
                            dtRelease.Columns["Vendor No"].ReadOnly = true;
                            dtRelease.Columns["PO #"].ReadOnly = true;
                            dtRelease.Columns["PO Line"].ReadOnly = true;
                            dtRelease.Columns["Release #"].ReadOnly = true;
                            dtRelease.Columns["Part Num"].ReadOnly = true;
                            dtRelease.Columns["Carrier"].ReadOnly = true;
                            dtRelease.Columns["Due Quantity"].ReadOnly = true;
                        //  dtRelease.Columns["Carrier_code"].ReadOnly = true;
                            dtRelease.Columns["Quantity"].ReadOnly = true;

                            this.vendorGrid.DataSource = bindingSourceRelease;
                            
                            vendorGrid.DisplayLayout.Bands[0].Columns["Quantity"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Due Quantity"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
                            vendorGrid.DisplayLayout.Bands[0].Columns["PO Line"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Tracking"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left;

                            vendorGrid.DisplayLayout.Bands[0].Columns["Due Quantity"].MaskInput = "n,nnn,nnn";
                            vendorGrid.DisplayLayout.Bands[0].Columns["Due Quantity"].MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Due Quantity"].MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeBoth;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Due Quantity"].MaskClipMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals;
                            this.vendorGrid.DisplayLayout.Bands[0].Columns["Due Quantity"].PromptChar = ' ';

                            vendorGrid.DisplayLayout.Bands[0].Columns["Quantity"].MaskInput = "n,nnn,nnn";
                            vendorGrid.DisplayLayout.Bands[0].Columns["Quantity"].MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Quantity"].MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeBoth;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Quantity"].MaskClipMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals;
                            this.vendorGrid.DisplayLayout.Bands[0].Columns["Quantity"].PromptChar = ' ';                            
                            vendorGrid.DisplayLayout.Bands[0].Columns["Vendor No"].Width = 150;
                            vendorGrid.DisplayLayout.Bands[0].Columns["PO #"].Width = 100;
                            vendorGrid.DisplayLayout.Bands[0].Columns["PO Line"].Width = 80;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Release #"].Width = 100;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Part Num"].Width = 100;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Carrier"].Width = 200;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Due Quantity"].Width = 100;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Quantity"].Width = 100;
                            vendorGrid.DisplayLayout.Bands[0].Columns["Tracking"].Width = 400;

                        }
                        catch (System.Exception exp)
                        {
                            //MessageBox.Show(exp.ToString());
                        }
                    }
               // }
            }
        }
        public DataSet getDataSet(string sqlquery, string constr)
        {
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlquery, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;
        }

        private void ultraButton1_Click(object sender, EventArgs e)
        {
            clearALL();
         
        }

        private void VendorTrackingNum_Load(object sender, EventArgs e)
        {
            try
            {
                // TODO: This line of code loads data into the 'dataSet4.DataTable1' table. You can move, or remove it, as needed.
                this.dataTable1TableAdapter.Fill(this.dataSet4.DataTable1);
                // TODO: This line of code loads data into the 'dataSet4.releases' table. You can move, or remove it, as needed.
                this.releasesTableAdapter.Fill(this.dataSet4.releases);
            }
            catch
            {
            }
          
        }

        private void releasesBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {   
            string sqlUpdate="";

            for (int i = 0; i < vendorGrid.Rows.Count; i++)
            {             
                string tracking_num = vendorGrid.Rows[i].Cells["Tracking"].Text.ToString();
                int row_id = Convert.ToInt32(vendorGrid.Rows[i].Cells["Release #"].Text.ToString());
                sqlUpdate = sqlUpdate + "update releases set m_tracking='" + tracking_num + "' where row_id=" + row_id;
            }
                SqlConnection cn = new SqlConnection(this.constrEpicor);
                cn.Open();
                SqlCommand cmd = new SqlCommand(sqlUpdate, cn);
                SqlTransaction tran = cn.BeginTransaction("Transaction");
                cmd.Transaction = tran;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                try
                {
                    da.Fill(ds);
                    tran.Commit();
                    MessageBox.Show("Updated Successfully");
                }
                catch
                {
                    tran.Rollback();
                    MessageBox.Show("Update failed");
                }
                cn.Close();
                bindgrid();
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            int row_id = Convert.ToInt32(vendorGrid.Selected.Rows[0].Cells[3].Value.ToString());
            string sqldelete=" delete from releases where row_id=" +row_id + " select @@error ";
            DataSet dsDelete=  getDataSet(sqldelete, this.constrEpicor);
            if (dsDelete.Tables.Count>0 )
                if (dsDelete.Tables[0].Rows.Count>0)
                    if (dsDelete.Tables[0].Rows[0][0].ToString()=="0" )
                        MessageBox.Show("Item Deleted successfully");

            bindgrid();
        }

        private void bindingSourceRelease_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            BindingSource bd = this.vendorGrid.DataSource as BindingSource;
            DataTable dt=bd.DataSource as DataTable;
            string filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\VendorTrackingNum.csv";
            ExportToCSV(dt, filename);

            MessageBox.Show("File has been exported to your desktop " + filename);
        }

        static void OpenCSVWithExcel(string path)
        {
            var ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Workbooks.OpenText(path, Comma: true);
            ExcelApp.Visible = true;
        }

        private void ExportToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers  
            //  sw.Write("ID");
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }  

    }
}
