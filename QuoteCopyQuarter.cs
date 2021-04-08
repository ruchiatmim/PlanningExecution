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
using Infragistics.Win;


namespace Version3
{
    public partial class frmQuoteCopy : Form
    {
        public frmQuoteCopy()
        {
            InitializeComponent();
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            string constr = ConfigurationManager.ConnectionStrings["epConnectionString"].ConnectionString;
            string qry = "insert into [vendor_sku]([sku_no],[last_recv_date] ,[vendor_no] ,[vend_sku] ,[note] ,[qty]  ,[curr_key] ,[m_quote_date]    ,m_qrt) (SELECT [sku_no],getdate() ,[vendor_no] ,[vend_sku]  ,[note] ,[qty]  ,[curr_key] ,[m_quote_date]    ,'" + cmbQtrTo.Value + "'  FROM [MIMDIST].[dbo].[vendor_sku]  where m_qrt= '" + this.cmbQtrFrom.Value + "')";
            string qryDelete = "delete from   [MIMDIST].[dbo].[vendor_sku]  where m_qrt= '" + cmbQtrTo.Value+ "'";
           
            qry = qryDelete + qry;                              //          +"  COMMIT ";

            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlTransaction tran = cn.BeginTransaction("Transaction");
            SqlCommand cmd = new SqlCommand(qry, cn);
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

        public void bindFromQrterFilter()
        {
            //  string sql = " SELECT 'Q'+DATENAME(Quarter, GETDATE())+'-'+cast(YEAR(getdate()) as varchar) as Quarter,1 as ord union SELECT 'Q'+DATENAME(Quarter,DATEADD(QQ,1, GETDATE()))+'-'+ cast(YEAR(DATEADD(QQ,1, GETDATE())) as varchar) as Quarter ,2 as ord  union SELECT 'Q'+DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))+'-'+ cast(YEAR(DATEADD(QQ,2, GETDATE())) as varchar) as Quarter,3 as ord  order by  ord ";
            string sql = "  SELECT cast(YEAR(getdate()) as varchar)+'Q'+replicate('0', (2-len(DATENAME(Quarter, GETDATE()))))+ DATENAME(Quarter, GETDATE()) as Quarter   union  SELECT cast(YEAR(DATEADD(QQ,-1, GETDATE())) as varchar) +'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,-1, GETDATE())))))+DATENAME(Quarter,DATEADD(QQ,-1, GETDATE())) as Quarter    union   SELECT cast(YEAR(DATEADD(QQ,-2, GETDATE())) as varchar)+'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,-2, GETDATE()))))) +DATENAME(Quarter,DATEADD(QQ,-2, GETDATE()))  as Quarter    ";
            DataTable dtQuarter = getDataSet(sql).Tables[0];
            DataRow drQtr = dtQuarter.NewRow();
            drQtr[0] = "";
            dtQuarter.Rows.InsertAt(drQtr, 0);
            cmbQtrFrom.DataSource = dtQuarter;
            cmbQtrFrom.DisplayMember = "Quarter";
            cmbQtrFrom.ValueMember = "Quarter";
            cmbQtrFrom.Value = dtQuarter.Rows[3][0].ToString();
        }

        public void bindToQrterFilter()
        {
            //  string sql = " SELECT 'Q'+DATENAME(Quarter, GETDATE())+'-'+cast(YEAR(getdate()) as varchar) as Quarter,1 as ord union SELECT 'Q'+DATENAME(Quarter,DATEADD(QQ,1, GETDATE()))+'-'+ cast(YEAR(DATEADD(QQ,1, GETDATE())) as varchar) as Quarter ,2 as ord  union SELECT 'Q'+DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))+'-'+ cast(YEAR(DATEADD(QQ,2, GETDATE())) as varchar) as Quarter,3 as ord  order by  ord ";
            string sql = "   SELECT cast(YEAR(DATEADD(QQ,1, GETDATE())) as varchar) +'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,1, GETDATE())))))+DATENAME(Quarter,DATEADD(QQ,1, GETDATE())) as Quarter    union   SELECT cast(YEAR(DATEADD(QQ,2, GETDATE())) as varchar)+'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))))) +DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))  as Quarter  union   SELECT cast(YEAR(DATEADD(QQ,3, GETDATE())) as varchar)+'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,3, GETDATE()))))) +DATENAME(Quarter,DATEADD(QQ,3, GETDATE()))  as Quarter  union SELECT cast(YEAR(DATEADD(QQ,4, GETDATE())) as varchar) +'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,4, GETDATE())))))+DATENAME(Quarter,DATEADD(QQ,4, GETDATE())) as Quarter   ";
            DataTable dtQuarter = getDataSet(sql).Tables[0];
             DataRow drQtr = dtQuarter.NewRow();
            drQtr[0] = "";
            dtQuarter.Rows.InsertAt(drQtr, 0);
            cmbQtrTo.DataSource = dtQuarter;
            cmbQtrTo.DisplayMember = "Quarter";
            cmbQtrTo.ValueMember = "Quarter";
           
            cmbQtrTo.Text= dtQuarter.Rows[1][0].ToString();
        }
        public DataSet getDataSet(string sqlStr)
        {
            string constr = ConfigurationManager.ConnectionStrings["epConnectionString"].ConnectionString;
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlStr, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;
        }

        private void frmQuoteCopy_Load(object sender, EventArgs e)
        {
            bindToQrterFilter();
            bindFromQrterFilter();
           
        }

        private void btnFilter_Click(object sender, EventArgs e)
        {

            string constr = ConfigurationManager.ConnectionStrings["epConnectionString"].ConnectionString;
            string qrySelect = "SELECT     sku_no AS [MIM PARTNM], vendor_no AS [Vend Code], vend_sku AS [MFG PART NUM], m_quote_date AS [Quote Date], ISNULL(curr_key, '') AS [CURRENCY], m_qrt AS [Quote QTR], CASE WHEN vendor_sku.vendor_no LIKE ('DELTA%')  THEN 'DELTA' ELSE vendor_sku.vendor_no END AS [Common Vend Name] FROM         dbo.vendor_sku    where m_qrt= '" + this.cmbQtrFrom.Value + "'";       
            
             
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(qrySelect, cn);
           
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                ugData.DataSource = dt;
              
            }
            catch (Exception exp)
            {
                string error = exp.Message.ToString();                
                MessageBox.Show("Failed to Update. Please check the data");
            }
            cn.Close();
            ugData.DisplayLayout.Bands[0].Override.AllowUpdate = DefaultableBoolean.False;
         
           

        }

        private void ultraButton1_Click(object sender, EventArgs e)
        {
            string constr = ConfigurationManager.ConnectionStrings["epConnectionString"].ConnectionString;
            string qryVendorSplit = " Exec Quote_Update_VendSplit_sp ";
                           
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(qryVendorSplit, cn);
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                cmd.ExecuteNonQuery();

            }
            catch (Exception exp)
            {
                string error = exp.Message.ToString();
                MessageBox.Show("Exception in copy vendor split");
            }
            cn.Close();
        }

      
    }
}
