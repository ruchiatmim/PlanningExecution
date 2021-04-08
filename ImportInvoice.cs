using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace Version3
{
    public partial class ImportInvoice : Form
    {
        public ImportInvoice()
        {
            InitializeComponent();
        }
        string constr = "";
        private void btnFilter_Click(object sender, EventArgs e)
        {
            string constr="";
            string site = cmbSite.SelectedText;

            string sqlQry = " SELECT Cast('false' as bit) as Chk   ,[InvNum]      ,[Dept]      ,[InvAmount]  FROM [dbo].[m_CSVUIView_rev] where yearWWeek='" + txtPeriod.Text.Trim() + "' and company='" + site + "'";
            DataTable dtInvoice = getDataTable(sqlQry);
            grdInvoice.DataSource = dtInvoice;
           

        }

        public DataTable getDataTable(string sqlQry)
        {

            string site = cmbSite.SelectedText;
           
            constr = ConfigurationManager.ConnectionStrings["mvFinanceConnectionString"].ConnectionString;   
            DataTable dtInvoice = new DataTable();

            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlQry, cn);
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);               
                da.Fill(dtInvoice);
                //  dtInvoice.Columns.Add(new DataColumn("inv", typeof(bool)));             
                //  tran.Commit();
            }
            catch (Exception exp)
            {
                string error = exp.Message.ToString();
                // tran.Rollback();

            }

            cn.Close();
            return dtInvoice;
        }

        private void ImportInvoice_Load(object sender, EventArgs e)
        {
            loadSite();
        }

        public void loadSite()
        {
             DataTable dtSite = new DataTable();
            dtSite.Columns.Add(new DataColumn("Site"));
        //    dtSite.Columns.Add(new DataColumn("Value"));
            DataRow dr = dtSite.NewRow();
            dr[0] = "MIMUS";
         
            dtSite.Rows.Add(dr);
            dr = dtSite.NewRow();
            dr[0] = "MIMNL";
         //   dr[1] = "AT";
            dtSite.Rows.Add(dr);
            cmbSite.DataSource = dtSite;
            cmbSite.DisplayMember = "Site";
            cmbSite.ValueMember = "Site";
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
          
            DataTable dt = grdInvoice.DataSource as DataTable;
        
            string Selectedinv ="'";
            string invName = "";
            foreach (DataRow dr in dt.Rows)
            {
                   if (dr[0].ToString()=="True") 
                Selectedinv = Selectedinv + dr[1] + "','";
               
            }
           String[] ArrInv = Selectedinv.Split(',');
           invName = ArrInv[0] + "_" + ArrInv[ArrInv.Length - 2];

           string site = cmbSite.SelectedText;
           string siteKey = "";
           if (site == "MIMNL")
               siteKey = "OF";
           else
               siteKey = "1";

           string sqlQry = "SELECT *  FROM [dbo].[ARCustExcelDownload] where [INVOICE NO.] in (" + Selectedinv.Substring(0, Selectedinv.Length - 2) + ") and [COMPANY CODE]='" + siteKey + "'  ORDER BY [INVOICE NO.], [LINE TYPE]";

            DataTable dtExport = getDataTable(sqlQry);
          
           
            
            if (site == "MIMNL")
                 siteKey = "NL";               
            else
                  siteKey = "US";
            
            string strFilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + txtPeriod.Text.Trim() + "_" + siteKey + "_" + invName.Substring(0,invName.Length-1) + ".csv";
            dtExport.ToCSV(strFilePath);
            MessageBox.Show("File has been exported to your desktop");
          
        }

        private void grdInvoice_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
            e.Layout.Bands[0].Columns[0].Width = 50;
            e.Layout.Bands[0].Columns[1].Width = 100;
            e.Layout.Bands[0].Columns[2].Width = 400;
            e.Layout.Bands[0].Columns[3].Width = 150;
        }

        private void chkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
           
            if (chkSelectAll.Checked)
            {

                foreach (var r in grdInvoice.Rows)
                {
                    r.Cells[0].Value = 1;
                }
            }
            else
            {
                foreach (var r in grdInvoice.Rows)
                {
                    r.Cells[0].Value = 0;
                }
            }
        }       
    
    }
}
