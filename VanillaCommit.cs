using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading.Tasks;
using Google.Apis.Download;
using System.Diagnostics;

using Google.Apis.Util;
using System.Net;
using System.Threading;
using System.Data.SqlClient;
using System.Configuration;

namespace Version3
{
    public partial class VanillaCommit : Form
    {

        string constr = ConfigurationManager.ConnectionStrings["MVIntConnectionString"].ConnectionString;
        public VanillaCommit()
        {
            InitializeComponent();
        }

       
      
       

     
        private void ultraButton1_Click(object sender, EventArgs e)
        {
            string sqlSave = "exec msdb.dbo.sp_start_job @job_name = 'YourJobName'";
            
                SqlConnection cn = new SqlConnection(constr);
                try
                {
                    cn.Open();
                    SqlCommand cmd = new SqlCommand(sqlSave, cn);
                    cmd.CommandTimeout = 0;
                    cmd.ExecuteNonQuery();
                    /*  SqlDataAdapter da = new SqlDataAdapter(cmd);
                      DataSet ds = new DataSet();
                      da.Fill(ds);
                      if (ds.Tables.Count > 0)
                          if (ds.Tables[ds.Tables.Count - 1].Rows.Count > 0)
                              if (ds.Tables[ds.Tables.Count - 1].Rows[0][0].ToString() == "0")
                                  return " Completed  ";*/

                }
                catch (Exception ex)
                {
                    MessageBox.Show("There is some error to save. Please check." + ex.Message.ToString());
                }
                finally
                {

                    cn.Close();
                }
        }

      

    }
}