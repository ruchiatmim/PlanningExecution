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
     
    public partial class MRP_Process : Form
    {

        string constr = ConfigurationManager.ConnectionStrings["mvEpicorConnectionString"].ConnectionString;
         string constrMvErp = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;

        public MRP_Process()
        {
            InitializeComponent();
            getProcess();
        }

      

        private void btnRunTempMRP_Click(object sender, EventArgs e)
        {
            lblTempProcessStatus.Text = "Running.....";
            string runProcess = "exec MRP_1_RUNTEMPSTEPS_SP  select @@Error  ";
            string Status = processMRP(runProcess);
            lblTempProcessStatus.Text = Status + System.DateTime.Now.ToString();
        }

        private string  processMRP(string sqlSave)
        {
            try
            {              
                SqlConnection cn = new SqlConnection(constrMvErp);
                cn.Open();
                SqlCommand cmd = new SqlCommand(sqlSave, cn);
                cmd.CommandTimeout = 0;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                if (ds.Tables.Count > 0)
                    if (ds.Tables[ds.Tables.Count - 1].Rows.Count > 0)
                        if (ds.Tables[ds.Tables.Count-1].Rows[0][0].ToString() == "0")
                            return " Completed  ";
                cn.Close();                
            }
            catch (Exception ex)
            {             
                MessageBox.Show("There is some error to save. Please check." + ex.Message.ToString());
            }            
            return " Failed  ";
        }

       /* private void btnRunLiveMRP_Click(object sender, EventArgs e)
        {
            lblLiveProcessStatus.Text = "Running.....";
            string runProcess = "exec MRP_2_TEMP_TO_LIVE_Table_SP select @@Error  ";
            string Status = processMRP(runProcess);
            lblLiveProcessStatus.Text = Status + System.DateTime.Now.ToString();
        }
        */
        private void getProcess()
        {
            string sqlProcess = "SELECT TOP 1000 [ID]       ,[ProcessName]      ,[ProcessDate]      ,[FileName]  FROM [PRODDIST].[dbo].[BuildPlanLoadProcessSteps] ";
            try
            {
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlCommand cmd = new SqlCommand(sqlProcess, cn);
                cmd.CommandTimeout = 0;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                int x =5;
                int y = 20;

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    Label lblProcessName = new Label();
                    lblProcessName.Name = "lbl" + ds.Tables[0].Rows[i]["ProcessName"].ToString();
                    lblProcessName.Text = ds.Tables[0].Rows[i]["ProcessName"].ToString();
                    lblProcessName.Location=new Point(x, y);
                    lblProcessName.Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 20.25F, FontStyle.Regular);
                    lblProcessName.Height = 31;
                    lblProcessName.Width = 425;
                 
                    Label lblProcessDate = new Label();
                    lblProcessDate.Name = "lbl" + ds.Tables[0].Rows[i]["ProcessDate"].ToString();
                    lblProcessDate.Text = "|     " + ds.Tables[0].Rows[i]["ProcessDate"].ToString();
                    lblProcessDate.Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 20.25F, FontStyle.Regular);
                    lblProcessDate.Location = new Point(425, y);
                    lblProcessDate.Height = 31;
                    lblProcessDate.Width = 425;

                    if (i % 2 == 0)
                    {
                        lblProcessName.BackColor = Color.Gray;
                        lblProcessDate.BackColor = Color.Gray;
                    }
 
                    panel1.Controls.Add(lblProcessName);
                    panel1.Controls.Add(lblProcessDate);
                    y = y + 40;

                }
                DataRow[] dr = ds.Tables[0].Select("ProcessName='LIVE MRP'");
                if (dr.Length>0)
                    lblTempProcessStatus.Text = " Completed " + dr[0]["ProcessDate"].ToString();

                dr = ds.Tables[0].Select("ProcessName='CLEAR TO FIRM'");
                if (dr.Length > 0)
                    lblCTFStatus.Text = " Completed " + dr[0]["ProcessDate"].ToString();
             
               
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("There is some error to save. Please check." + ex.Message.ToString());
            }

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.lblCTFStatus.Text = "Running.....";
            string runProcess = "exec [dbo].[MRP_CTF_Table_SP] select @@Error  ";
            string Status = processMRP(runProcess);
            lblCTFStatus.Text = Status + System.DateTime.Now.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string runProcess = "exec mvquality.master.dbo.sp_start_sql_agent_job_wait 'Run CTF (Manual)'";

                try
            {              
                SqlConnection cn = new SqlConnection(constrMvErp);
                cn.Open();
                SqlCommand cmd = new SqlCommand(runProcess, cn);
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();
                MessageBox.Show("CTF Data has been exported to BQ. ");
                cn.Close();                
            }
            catch (Exception ex)
            {             
                MessageBox.Show("Error to Start the job. Please contact to administrator");
            }            
           
        }

        private void MRP_Process_Load(object sender, EventArgs e)
        {

        }
    }
}
