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

     
    public partial class BTS_Netgear : Form
    {

        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        public BTS_Netgear()
        {
            InitializeComponent();
        }

        private void BTS_Netgear_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'bTS.BTS_GoogleIntInvTranNetgear' table. You can move, or remove it, as needed.
            this.bTS_GoogleIntInvTranNetgearTableAdapter.Fill(this.bTS.BTS_GoogleIntInvTranNetgear);
            this.txtStartdate.Value = DateTime.Now.AddDays(-2);
            loadStatus();
        }
        public void loadStatus()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Status");
            DataRow dr = dt.NewRow();
            dr["Status"] = "A";
            dt.Rows.Add(dr);
            
            dr = dt.NewRow();
            dr["Status"] = "N";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Status"] = "X";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Status"] = "Y";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Status"] = "ALL";
            dt.Rows.Add(dr);
            cmbStatus.DataSource = dt;
            ddStatus.DataSource = dt;
            cmbStatus.SelectedText = "ALL";

        }
        public void getSerialNum(string Ident)
        {
            //WebDataGridSerial.DataSource = null;
            string sql = " SELECT [auto_id],status,[serial_num],[parent_serial_num] ,[asset_tag_num],reference  FROM [MIMDIST].[dbo].[BTS_GoogleIntInvTranSerial]  where uniqueIdent='" + Ident + "'";
            SqlConnection cn = new SqlConnection(constr);
            SqlCommand cmd = new SqlCommand(sql, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet dsSerial = new DataSet();
            da.Fill(dsSerial);
            ugBTSSerial.DataSource = dsSerial.Tables[0];
        
        }


        string uniqueIdent;
        private void ugBTS_AfterSelectChange(object sender, Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs e)
        {
            if (this.ugBTS.Selected.Rows.Count > 0)
            {
                uniqueIdent = ugBTS.Selected.Rows[0].Cells["uniqueIdent"].Value.ToString();
            }
            txtUniqueIdentifier.Text = uniqueIdent;
            getSerialNum(uniqueIdent);
        }

        private void ubDelete_Click(object sender, EventArgs e)
        {
            int serial_auto_id = 0;
            if (this.ugBTSSerial.Selected.Rows.Count > 0)
            {
                if (this.ugBTSSerial.Selected.Rows[0].Cells[0].Text.ToString() != "")
                {
                    serial_auto_id = Convert.ToInt32(this.ugBTSSerial.Selected.Rows[0].Cells[0].Text.ToString());
                    string deleteQry = "delete from BTS_GoogleIntInvTranSerial where auto_id=" + serial_auto_id + "  select @@error ";
                    SqlConnection cn = new SqlConnection(constr);
                    cn.Open();
                    SqlCommand cmd = new SqlCommand(deleteQry, cn);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    cn.Close();
                    if (ds != null)
                        if (ds.Tables.Count > 0)
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                if (ds.Tables[0].Rows[0][0].ToString() == "0")
                                    MessageBox.Show("Serial Num has been deleted Successfully");
                            }
                }
            }
        }

        private void ubSave_Click(object sender, EventArgs e)
        {
            string qry = "";
            //MessageBox.Show(vendorGrid.Rows.Count.ToString());
            for (int i = 0; i < ugBTSSerial.Rows.Count; i++)
            {
                int auto_id = 0;
                string serial_num = ugBTSSerial.Rows[i].Cells["serial_num"].Value.ToString().Trim();

                string parent_serial_num = ugBTSSerial.Rows[i].Cells["parent_serial_num"].Value.ToString().Trim();
                string asset_tag_num = ugBTSSerial.Rows[i].Cells["asset_tag_num"].Value.ToString().Trim();

                if ((ugBTSSerial.Rows[i].Cells["auto_id"].Value != null) && (ugBTSSerial.Rows[i].Cells["auto_id"].Value.ToString().Trim() != ""))
                    auto_id = Convert.ToInt32(ugBTSSerial.Rows[i].Cells["auto_id"].Value.ToString().Trim());
                string app = "";
                if (serial_num != "")
                    app = "NG";
                else
                    app = "EPI";
                if (auto_id == 0)
                {
                    qry = qry + " insert into BTS_GoogleIntInvTranSerial(ProcessDate,app,asset_tag_num,epi_tran_type,fileDate,filename,folder,location,parent_serial_num,part_no,reference,region,serial_num,status,tran_date,tran_id,uniqueIdent) select ProcessDate,'"+app+"','" + asset_tag_num + "',epi_tran_type,fileDate, filename,folder,location,'" + parent_serial_num + "',part_no,reference,region,'" + serial_num + "','N',tran_date,tran_id,uniqueIdent  from BTS_GoogleIntInvTranNetgear where uniqueIdent='" + uniqueIdent + "'";
                }
                else
                {
                    qry = qry + "   update dbo.BTS_GoogleIntInvTranSerial  set asset_tag_num='" + asset_tag_num + "', parent_serial_num='" + parent_serial_num + "', serial_num='" + serial_num + "' where auto_id=" + auto_id.ToString();
                }

            }
            //   string constr = ConfigurationManager.ConnectionStrings["epConnectionStringTest"].ConnectionString;

            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(qry, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            getSerialNum(uniqueIdent);
        }

        private void ubAddNew_Click(object sender, EventArgs e)
        {
            this.ugBTSSerial.DisplayLayout.Bands[0].AddNew();
        }

        private void ultraButton1_Click(object sender, EventArgs e)
        {
            string sql = "select  *  from BTS_GoogleIntInvTranNetgear where  uniqueIdent is not null  ";  //,replace(filename,'\\\\mimsftp\\e$\\sftp\\msgGoogle\\PLATFORM\\PROD\\','') as file1


            if (cmbStatus.Text != "ALL")
            {
                sql = sql + " and   status='" + cmbStatus.Text + "'";
            }

            if (txtTranID.Text.Trim() != "")
                sql = sql + " and   tran_id=" + txtTranID.Text.Trim();

            if (this.utUniqueIdent.Text.Trim() != "")
                sql = sql + " and   uniqueIdent='" + utUniqueIdent.Text.Trim() + "'";

            if (this.txtStartdate.Value != "" && this.txtEndDate.Value != "")
            {

                DateTime st_date = Convert.ToDateTime(Convert.ToDateTime(this.txtStartdate.Value).ToShortDateString());
                DateTime end_date = Convert.ToDateTime(Convert.ToDateTime(this.txtEndDate.Value).ToShortDateString());
                sql = sql + " and   tran_date>='" + st_date + "'  and tran_date<='" + end_date + "'  ";
             //   sql = sql + " and   tran_date>='" + this.txtStartdate.Value + "'  and tran_date<='" + this.txtEndDate.Value + "'  ";
            }
            else
            {
                MessageBox.Show("Please select start tran date and end tran date");
            }
            
            sql = sql + "  order by auto_id desc";
            //   SqlDataSource1.SelectCommand = sql;

            SqlConnection cn = new SqlConnection(constr);
            SqlCommand cmd = new SqlCommand(sql, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            ugBTS.DataSource = ds.Tables[0];
            if (ugBTS.Rows.Count > 0)
            {
                uniqueIdent = ugBTS.Rows[0].Cells["uniqueIdent"].Value.ToString();
            }
            txtUniqueIdentifier.Text = uniqueIdent;
            getSerialNum(uniqueIdent);
        }
        private void ugBTS_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
        {
            int auto_id = 0;
            string reference = "";
            string status = "";
            if (e.Cell.Column.Index == 1 || e.Cell.Column.Index == 20)
            {

                //  if (this.ugBTS.Selected.Rows.Count > 0)
                //  {
                auto_id = Convert.ToInt32(e.Cell.Row.Cells["auto_id"].Value.ToString());
                status = e.Cell.Row.Cells["status"].Value.ToString();
                reference = e.Cell.Row.Cells["reference"].Value.ToString();
                string qry = "update dbo.BTS_GoogleIntInvTranNetgear set status='" + status + "', reference='" + reference + "'   where auto_id=" + auto_id;
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlCommand cmd = new SqlCommand(qry, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                cn.Close();
                // }
            }
       

        }

        private void ugBTS_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {

        }
          
    }
}
