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
    public partial class kb_PopupPartAssign : Form
    {
         String siteCode ;
    
      //   string constrTEST = "Data Source=mverp;Initial Catalog=testdist;user id=sa;password=mimi~100;";
         string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
         public kb_PopupPartAssign(String KBLOc, String partnum, String site, int bin_qty, string contType)
        {
            InitializeComponent();
            txtPartNum.Text = partnum;
            lblKBLoc.Text = KBLOc;
            siteCode = site;
            lblBinQTY.Text = bin_qty.ToString();
            lblContType.Text = contType;

        }
             

        private void Button3_Click(object sender, EventArgs e)
        {
            updatePartNum();
        }

      
    
    private void Button2_Click(System.Object sender  , System.EventArgs e )
    {
        txtPartNum.Text = "";
    }

 private void PictureBox1_Click(System.Object sender  , System.EventArgs e ) 
 {
      

      

 }

    public void updatePartNum()
    {

        String sqlstr  = siteCode.ToString();

        if (txtPartNum.Text =="") 
            sqlstr = "Delete from [KanBanLayoutPart] where kb_loc='" + lblKBLoc.Text.Trim() + "' select @@rowcount";
        else
            sqlstr = "if exists (select 1 from  [KanBanLayoutPart] where  kb_loc='" + lblKBLoc.Text.Trim() + "')    update [KanBanLayoutPart] set  partnum='" + txtPartNum.Text.Trim() + "'  where kb_loc='" + lblKBLoc.Text.Trim() + "' else insert into [KanBanLayoutPart](SiteCode,partnum,kb_loc) values('" + siteCode.ToString() + "','" + txtPartNum.Text.Trim() + "','" + lblKBLoc.Text.Trim() + "')  select @@rowcount";


        SqlConnection conn = new System.Data.SqlClient.SqlConnection(constr);
        conn.Open();
        SqlCommand cmd =new SqlCommand(sqlstr, conn);
        cmd.CommandTimeout = 0;
        SqlDataAdapter da =new SqlDataAdapter(cmd);
        DataSet dsKanbanLayoutUpdate =new DataSet();

        da.Fill(dsKanbanLayoutUpdate);
        int rowcount  = 0;
        rowcount = Convert.ToInt32(dsKanbanLayoutUpdate.Tables[0].Rows[0][0].ToString());

        if (rowcount > 0)
        {
            // MessageBox.Show("Updated successfully");
        }
        else
            MessageBox.Show("Update fails");
        
        this.Close();
    }

  

    private void label3_Click(object sender, EventArgs e)
    {
        if (DialogResult.Yes == MessageBox.Show("Are you sure you want to unassign", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
        {
            txtPartNum.Text = "";
        }

    }

   

    }
}
