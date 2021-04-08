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
    public partial class NotePad : Form
    {
        public string notes = "";
        public string var_PO = "";
        
        public NotePad( string note,string PO)
        {
            InitializeComponent();
            Note2.Text = note;
            notes = note;
            var_PO = PO;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
           
            notes = Note2.Text;
            if (notes.Length > 255)
            {
                notes = notes.Substring(0, 254);
            }
            string qry = "";
            if (this.Text == "INTERNAL NOTES")
                qry = " update purchase set user_def_fld4='" + notes + "'  where po_no = " + var_PO;
            else if (this.Text == "EXTERNAL NOTES")
                qry = " update purchase set note='" + notes + "'  where po_no = " + var_PO;
            if (this.Text == "INTERNAL NOTES" || (this.Text == "EXTERNAL NOTES"))
            {
                string constr = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;
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
                    //  MessageBox.Show("Updated successfully");
                }
                catch (System.Exception exp)
                {
                    string error = exp.Message.ToString();
                    tran.Rollback();
                    MessageBox.Show("Failed to Update:" + error);
                }
                cn.Close();
            }

           this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
