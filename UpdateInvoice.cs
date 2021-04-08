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
    public partial class frmUpdateInvoice : Form
    {
        string constrMVERP = ConfigurationManager.ConnectionStrings["mverpmaterialinmotionincConnectionString"].ConnectionString;
        public frmUpdateInvoice()
        {
            InitializeComponent();
        }

        private void ultraButton1_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            string qry="exec m_sp_m_InvoiceAudit_BTS select @@error";
            try
            {
                DataSet ds = getDataSet(qry);
                if (ds.Tables[0].Rows[0][0].ToString() == "0")
                    MessageBox.Show("BTS Data has been updated successfully");
            }
            catch(Exception exp)
            {
                string error = exp.Message.ToString();
                MessageBox.Show("Failed to update BTS Data. Please check!" + error);
            }
            Cursor.Current = Cursors.Default;
        }

        public DataSet getDataSet(string strqry)
        {
            SqlConnection cn = new SqlConnection(constrMVERP);
            cn.Open();
            SqlCommand cmd = new SqlCommand(strqry, cn);
            cmd.CommandTimeout = 0;
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;
        }

     
        private void btnUpdateSBASG_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            string qry = "exec m_sp_m_InvoiceAudit_SG select @@error";
            try
            {
                DataSet ds = getDataSet(qry);
                if (ds.Tables[0].Rows[0][0].ToString() == "0")
                    MessageBox.Show("BTS Data has been updated successfully");
            }
            catch (Exception exp)
            {
                string error = exp.Message.ToString();
                MessageBox.Show("Failed to update BTS Data. Please check!" + error);
            }
            Cursor.Current = Cursors.Default;
        }

        private void btnUpdateSBAUS_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            string qry = "exec m_sp_m_InvoiceAudit_AT select @@error";
            try
            {
                DataSet ds = getDataSet(qry);
                if (ds.Tables[0].Rows[0][0].ToString() == "0")
                    MessageBox.Show("BTS Data has been updated successfully");
            }
            catch (Exception exp)
            {
                string error = exp.Message.ToString();
                MessageBox.Show("Failed to update BTS Data. Please check!" + error);
            }
            Cursor.Current = Cursors.Default;

        }

        private void btnUpdateSBANL_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            string qry = "exec m_sp_m_InvoiceAudit_NL select @@error";
            try
            {
                DataSet ds = getDataSet(qry);
                string rr=ds.Tables[0].Rows[0][0].ToString();
                if (ds.Tables[0].Rows[0][0].ToString() == "0")
                    MessageBox.Show("BTS Data has been updated successfully");
            }
            catch (Exception exp)
            {
                string error = exp.Message.ToString();
                MessageBox.Show("Failed to update BTS Data. Please check!" + error);
            }
            Cursor.Current = Cursors.Default;
        }
    }
}
