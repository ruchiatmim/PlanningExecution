using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Configuration;

namespace Version3
{
    public partial class KB_PopupOrderAssign : Form
    {

        String sitecode;
        String kb_loc;
        int bin_qty ;
        string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
        string constrEpicor = ConfigurationManager.ConnectionStrings["epConnectionString"].ConnectionString;
       // string constrTest = "Data Source=mverp;Initial Catalog=testdist;user id=sa;password=mimi~100;";
        DataSet dsKanbanOrder;
        public KB_PopupOrderAssign(String kb_loc  , String sitecode, string partnum,int binqty)
        {
            InitializeComponent();             
            kb_loc = kb_loc;
            sitecode = sitecode;
            lblHeading.Text = kb_loc;
            lblPartNum.Text = partnum;
            bin_qty = binqty;
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(constr);
            conn.Open();
            String cmdtxt = "SELECT  [SiteCode],[PONum],[POLine],[OrderNum] ,REPLACE(part_no, '-R', '') as [PartNum],[KB_Loc] , [Aged]   FROM  [dbo].[v_KBLayout_OpenOrders]  where [SiteCode] ='" + sitecode + "'  and kb_loc='" + kb_loc + "' order by PartNum";
            SqlCommand cmd =new SqlCommand(cmdtxt, conn);
            cmd.CommandTimeout = 0;
            SqlDataAdapter da =new SqlDataAdapter(cmd);
            dsKanbanOrder=new DataSet();
            da.Fill(dsKanbanOrder);
            int xloc, yloc ;
            yloc = 25;
                      
            
                for (int i =0;i<bin_qty;i++)
                {
                    xloc = 10;
                    Label lblOrder=new Label();
                    lblOrder.Name = "lblOrder" + (i + 1).ToString();
                    lblOrder.Text = (i + 1).ToString() + " : ";
                    lblOrder.Width = 20;

                    lblOrder.Location=new Point(xloc, yloc);
                    xloc = xloc + 25;
                    TextBox txtOrder=new TextBox();
                    txtOrder.Name = "txtOrder" + (i + 1).ToString();
                    txtOrder.Width = 120;
                    txtOrder.Location = new Point(xloc, yloc);
                    txtOrder.MaxLength = 12;
                    txtOrder.Leave += new EventHandler(txtOrder_leave);
                    xloc = xloc + 120;

                    Label lblPartNo = new Label();
                    lblPartNo.Name = "lblPartNo" + (i + 1).ToString();
                  
                    lblPartNo.Width = 100;
                    lblPartNo.Location = new Point(xloc, yloc);

                    if (i < dsKanbanOrder.Tables[0].Rows.Count)
                    {
                        txtOrder.Text = dsKanbanOrder.Tables[0].Rows[i]["OrderNum"].ToString();
                        lblPartNo.Text = dsKanbanOrder.Tables[0].Rows[i]["PartNum"].ToString();
                        ord_list = ord_list + "," + dsKanbanOrder.Tables[0].Rows[i]["OrderNum"].ToString();
                    }
                    
                    
                    xloc = xloc + 100;
                                       
                    Label lblDelete = new Label();
                    lblDelete.Name = "lblDelete" + (i + 1).ToString();
                                
                    lblDelete.AutoSize = true;
                    lblDelete.BackColor = System.Drawing.Color.WhiteSmoke;
                    lblDelete.Font = new System.Drawing.Font("Wingdings 2", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
                    lblDelete.ForeColor = System.Drawing.Color.Crimson;
                    lblDelete.Location = new Point(xloc, yloc+2);
                    lblDelete.Size = new System.Drawing.Size(16, 14);
                    lblDelete.Text = "O";
                    lblDelete.Click += new System.EventHandler(Delete_Click);

                    yloc = yloc + 30;
                    pnlOrder.Controls.Add(lblOrder);
                    pnlOrder.Controls.Add(txtOrder);
                    pnlOrder.Controls.Add(lblPartNo);
                    pnlOrder.Controls.Add(lblDelete);       



            }
            }

      
   
        private void btnDone_Click(object sender, EventArgs e)
        {
            AssignKBBin();
        }

        private void AssignKBBin()
        {
            String po_list = "";
            bool dupOrder=false;
            for (int i = 1;i<= bin_qty;i++)
                {
                    TextBox txtorder   = pnlOrder.Controls.Find("txtOrder" + i.ToString(), false)[0] as TextBox;
                    if (txtorder.Text.Trim() != "")
                    {
                        if (po_list.Contains(txtorder.Text.Trim()))
                        {
                            dupOrder=true;
                            txtorder.BackColor=Color.Red;
                            txtorder.Focus();
                            goto gend ;
                        }
                        else
                          po_list = po_list +"'"+ txtorder.Text.Trim() + "',";
                        }
                }

           
            String strsql = "  update pur_list set [KBbin]=''   where [KBbin]='" + lblHeading.Text.Trim() + "'  ";
        if (po_list.Length > 1)
        {
            po_list = po_list.Substring(0, po_list.Length - 1);
            strsql += "  update pur_list set [KBbin]='" + lblHeading.Text.Trim() + "' where po_no in (select ponum from  mverp.mimdist.dbo.[v_KBLayout_OpenOrders] where ordernum in  (" + po_list + "))  ";
        }

        strsql += " select @@error";

        SqlConnection conn = new System.Data.SqlClient.SqlConnection(constrEpicor);
        conn.Open();

        SqlCommand cmd =new SqlCommand(strsql, conn);
        cmd.CommandTimeout = 0;
        SqlDataAdapter da =new SqlDataAdapter(cmd);
        DataSet dsKanbanOrder=new DataSet();
        da.Fill(dsKanbanOrder);
        if (dsKanbanOrder.Tables[0].Rows[0][0].ToString() != "0")
            MessageBox.Show("Fail to assign order", "Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
        else
            this.Close();

    gend:
            if  (dupOrder)
        MessageBox.Show("Duplicate ORder");

        }
         private void Delete_Click(System.Object sender  ,System.EventArgs e  )
         {
        
            if (DialogResult.Yes==MessageBox.Show("Are you sure you want to unassign", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                String ind = (sender as Label).Name.ToString().Replace("lblDelete", "");
                TextBox txt   = pnlOrder.Controls.Find("txtOrder" + ind, false)[0] as TextBox;
                txt.Text = "";
            }

         }
         string ord_list;
        private void txtOrder_leave(object sender, EventArgs e)
        {
            TextBox txt=sender as TextBox;
            string str=txt.Text.Trim();
            txt.BackColor = Color.White;
            if (str.Length>0)
                if (!Regex.IsMatch(str, "^[A-Za-z]{6}[0-9]{6}$"))
                {
                    MessageBox.Show("Valid order must have 6 letter prefix and order number ", "Invalid Order Number", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txt.Text = "";
                    txt.Focus();

                }
            if (txt.Text.Trim() != "")
            {
                Label lblPart = pnlOrder.Controls.Find("lblPartNo" + txt.Name.Replace("txtOrder", ""), false)[0] as Label;
                if (dsKanbanOrder.Tables[0].Select("ordernum='" +txt.Text+"'").Length>0)
                    lblPart.Text = dsKanbanOrder.Tables[0].Select("ordernum='" + txt.Text + "'")[0]["PartNum"].ToString();
                if (lblPart.Text != lblPartNum.Text)
                    lblPart.ForeColor = Color.Red;
            }
          
        
        }
        private string getPartNum(string order_num)
        {
            
            string sqlPartNum = "select part_no from pur_list where po_no=" + order_num;
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(constrEpicor);
            conn.Open();

            SqlCommand cmd = new SqlCommand(sqlPartNum, conn);
            cmd.CommandTimeout = 0;
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet dsOrder = new DataSet();
            da.Fill(dsOrder);

            if (dsOrder !=null)
                if (dsOrder.Tables.Count>0)
                    if (dsOrder.Tables[0].Rows.Count>0)
                        return dsOrder.Tables[0].Rows[0]["part_no"].ToString();
            return "";
        }


    }
}
