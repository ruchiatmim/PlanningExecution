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

namespace Version3
{
    public partial class AlternatePartPopup : Form
    {
    
        public  AlternatePartPopup( int part_id)
        {
            InitializeComponent();
            txtPartID.Text=part_id.ToString();
            this.Text = " Part id :" + part_id.ToString();
            GetAllAltPartNum();
        }

        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        private void btnAdd_Click(object sender, EventArgs e)
        {
            int id = 0;
             string sqlStr = "";
            if (txtAltPartNum.Text.Trim()!="")
            {
           sqlStr = " insert into dbo.vc_vend_part_alt(vend_part_id,AltPartNum) values('"+txtPartID.Text.Trim()+"','"+txtAltPartNum.Text.Trim()+"')";
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            try
            {
                SqlCommand cmd = new SqlCommand(sqlStr, cn);
                cmd.ExecuteNonQuery();
                GetAllAltPartNum();
            }
            catch (Exception ex)
            {

            }
            }
     
        }


        public void GetAllAltPartNum()
        {
            string sqlStr = "select * from dbo.vc_vend_part_alt where vend_part_id= "+txtPartID.Text.Trim();
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlStr, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            listBox1.Items.Clear();
            foreach (DataRow dr in ds.Tables[0].Rows)
                listBox1.Items.Add(dr["AltPartNum"].ToString());
            clear();
        }
        public void clear()
        {
            txtAltPartNum.Text = "";
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            string part_num="";
            if (listBox1.SelectedIndex > -1)
            {
                part_num= listBox1.SelectedItem.ToString();
                DeleteAltPartNum(part_num);
            }
        }

        public void DeleteAltPartNum(string partnum)
        {
            string sqlStr = "Delete from dbo.vc_vend_part_alt where vend_part_id= " + txtPartID.Text.Trim() + " and [AltPartNum]='" + partnum + "'";
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlStr, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            GetAllAltPartNum();
        }

        public void existPart(string partnum)
        {
            string sqlStr = "select * from inv_master where part_num='" + partnum + "' and active='Y'";
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlStr, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            GetAllAltPartNum();
        }
    }
}
