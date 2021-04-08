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
    public partial class KanbanOrderAssign : Form
    {
        string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
        public KanbanOrderAssign()
        {
            InitializeComponent();
            assignEventToBin();
           
            showALLPart();
        }
        DataSet dsKanbanOrder;
        public void showALLPart()
        {
            int x = 10;
            int y = 10;


            SqlConnection conn = new System.Data.SqlClient.SqlConnection(constr);
            conn.Open();
            String cmdtxt = "SELECT *   FROM  [dbo].[v_KBLayout_OpenOrders]  where [SiteCode] ='AT' order by part_no  ";
            SqlCommand cmd = new SqlCommand(cmdtxt, conn);
            cmd.CommandTimeout = 0;
            SqlDataAdapter da = new SqlDataAdapter(cmd);
             dsKanbanOrder = new DataSet();
            da.Fill(dsKanbanOrder);

            DataTable uniqueCols = dsKanbanOrder.Tables[0].DefaultView.ToTable(true, "part_no");
            

            foreach (DataRow dr in uniqueCols.Rows)
            {
                String part = dr["part_no"].ToString();
                Label lblPart = new Label();
                lblPart.Text = part;
        
                byte bt = (byte)(0);
                lblPart.Font = new Font("Microsoft Sans Serif", 8.0F, FontStyle.Regular, GraphicsUnit.Point, ((System.Byte)(0)));
                lblPart.Click += new EventHandler(this.lblPart_Click);
                lblPart.Location = new System.Drawing.Point(x, y);
                y = y + 25;
                PnlPartNum.Controls.Add(lblPart);
            }

        }
        public void clearButtonColor()
        {
            foreach (Control ctrl in PnlPartNum.Controls)
                ctrl.BackColor = Color.Transparent;

        }


        private void assignDefaultButton()
        {
            int xloc, yloc;
            xloc = 10;
            yloc = 10;

            try
            {
                foreach (DataRow dr in this.kanbanLayoutSecition1.dsKanbanLayout.Tables[0].Rows)
                {
                    String kb_loc;
                    String PartNum;
                    kb_loc = dr["KB_Loc"].ToString();
                    PartNum = dr["Partnum"].ToString();
                    Button btnloc = new Button();
                    if (this.kanbanLayoutSecition1.PanelA.Controls.Find(kb_loc, false).Length > 0)
                        btnloc = this.kanbanLayoutSecition1.PanelA.Controls.Find(kb_loc, false)[0] as Button;
                    else if (this.kanbanLayoutSecition1.PanelB.Controls.Find(kb_loc, false).Length > 0)
                        btnloc = this.kanbanLayoutSecition1.PanelB.Controls.Find(kb_loc, false)[0] as Button;
                    else if (this.kanbanLayoutSecition1.PanelC.Controls.Find(kb_loc, false).Length > 0)
                        btnloc = this.kanbanLayoutSecition1.PanelC.Controls.Find(kb_loc, false)[0] as Button;
                    else if (this.kanbanLayoutSecition1.PanelD.Controls.Find(kb_loc, false).Length > 0)
                        btnloc = this.kanbanLayoutSecition1.PanelD.Controls.Find(kb_loc, false)[0] as Button;
                    else
                    {
                        Label LBL = new Label();
                        LBL.Text = kb_loc;
                        LBL.Location = new System.Drawing.Point(xloc, yloc);
                       // kanbanLayoutSecition1.pPanelNF.Controls.Add(LBL);
                        yloc = yloc + 30;

                    }

                    if (PartNum == "UNASSIGNED")
                    {
                        btnloc.BackColor = Color.DimGray;
                        btnloc.ForeColor = Color.White;
                    }
                    else
                    {
                        btnloc.BackColor = Color.Gainsboro;
                        btnloc.ForeColor = Color.Black;
                    }

                    ToolTip1.SetToolTip(btnloc, PartNum);
                }

            }

            catch (Exception ex)
            {

                //'If btnloc Is Nothing Then
                //'    Dim LBL As New Label
                //'    LBL.Text = kb_loc
                //'    PanelNF.Controls.Add(LBL)

                //'End If

            }

        }
        private void lblPart_Click(System.Object sender, System.EventArgs e)
        {
            clearButtonColor();
           
            Label lblpart = sender as Label;
            lblpart.BackColor = Color.Bisque;
            String part_num = lblpart.Text;
            assignDefaultButton();


            DataRow[] drPart = dsKanbanOrder.Tables[0].Select("part_no='" + part_num + "' and kb_loc<>''");
            foreach (DataRow drloc in drPart)
            {
                String kb_loc = drloc["kb_loc"].ToString();
                
                Button btnloc = new Button();
                
                if (this.kanbanLayoutSecition1.PanelA.Controls.Find(kb_loc, false).Length > 0)
                    btnloc = this.kanbanLayoutSecition1.PanelA.Controls.Find(kb_loc, false)[0] as Button;
                else if (this.kanbanLayoutSecition1.PanelB.Controls.Find(kb_loc, false).Length > 0)
                    btnloc = this.kanbanLayoutSecition1.PanelB.Controls.Find(kb_loc, false)[0] as Button;
                else if (this.kanbanLayoutSecition1.PanelC.Controls.Find(kb_loc, false).Length > 0)
                    btnloc = this.kanbanLayoutSecition1.PanelC.Controls.Find(kb_loc, false)[0] as Button;
                else if (this.kanbanLayoutSecition1.PanelD.Controls.Find(kb_loc, false).Length > 0)
                    btnloc = this.kanbanLayoutSecition1.PanelD.Controls.Find(kb_loc, false)[0] as Button;
                btnloc.BackColor = Color.DarkRed;
                btnloc.ForeColor = Color.Black;

            }

        }


        public void assignEventToBin()
        {

            foreach (Control ctrlbtn in this.kanbanLayoutSecition1.PanelA.Controls)
            {
                if (ctrlbtn is Button)
                {
                    ctrlbtn.Click += new EventHandler(showpart);
                  
                }
            }

            foreach (Control ctrlbtn in this.kanbanLayoutSecition1.PanelB.Controls)
            {
                if (ctrlbtn is Button)
                {
                    ctrlbtn.Click += new EventHandler(showpart);
                  
                }
            }

            foreach (Control ctrlbtn in this.kanbanLayoutSecition1.PanelC.Controls)
            {
                if (ctrlbtn is Button)
                {
                    ctrlbtn.Click += new EventHandler(showpart);
                  
                }
            }
            foreach (Control ctrlbtn in this.kanbanLayoutSecition1.PanelD.Controls)
            {
                if (ctrlbtn is Button)
                {
                    ctrlbtn.Click += new EventHandler(showpart);
                  
                }
            }
            // this.kanbanLayoutSecition1.BR005.Click += new EventHandler(showpart);
            //   this.kanbanLayoutSecition1.BR005.Click += new EventHandler(showpart);
            //  this.kanbanLayoutSecition1.BR005.Click += new EventHandler(showpart);

        }


    

        public void showpart(System.Object sender, System.EventArgs e)  //               this.kanbanLayoutSecition1.BR005.Click, BR007.Click, BR006.Click, RR018.Click, RR017.Click, RR016.Click, RR015.Click, RR014.Click, RR013.Click, RR012.Click, RR011.Click, RR010.Click, RR009.Click, RR008.Click, RR007.Click, RR006.Click, RR005.Click, RR004.Click, RR003.Click, RR002.Click, RR001.Click, BR066.Click, BR065.Click, BR064.Click, BR063.Click, BR062.Click, BR061.Click, BR060.Click, BR011.Click, BR010.Click, BR009.Click, BR008.Click, BP110.Click, BP109.Click, BP108.Click, BP107.Click, BP106.Click, BP105.Click, BP104.Click, BP103.Click, BP102.Click, BP101.Click, BP100.Click, BP099.Click, BP098.Click, BP097.Click, BP096.Click, BP095.Click, BP094.Click, BP093.Click, BP092.Click, BP091.Click, BP090.Click, BP089.Click, BP088.Click, BP087.Click, BP086.Click, BP085.Click, BP084.Click, BP083.Click, BP082.Click, BP081.Click, BP080.Click, BP079.Click, BP078.Click, BP077.Click, BP076.Click, BP075.Click, BP074.Click, BP073.Click, BP072.Click, BP071.Click, BP059.Click, BP058.Click, BP057.Click, BP056.Click, BP055.Click, BP054.Click, BP053.Click, BP052.Click, BP051.Click, BP050.Click, BP049.Click, BP048.Click, BP047.Click, BP046.Click, BP045.Click, BP044.Click, BP043.Click, BP042.Click, BP041.Click, BP040.Click, BP039.Click, BP038.Click, BP037.Click, BP034.Click, BP033.Click, BP032.Click, BP031.Click, BP030.Click, BP029.Click, BP028.Click, BP027.Click, BP026.Click, BP025.Click, BP024.Click, BP023.Click, BP022.Click, BP021.Click, BP020.Click, BP019.Click, BP018.Click, BP017.Click, BP016.Click, BP015.Click, BP014.Click, BP013.Click, BP012.Click, RP029.Click, RP028.Click, RP027.Click, RP026.Click, RP025.Click, RP024.Click, RP023.Click, RP022.Click, RP021.Click, RP020.Click, RP019.Click, RP018.Click, RP017.Click, RP016.Click, RP015.Click, RP014.Click, RP013.Click, RP012.Click, RP011.Click, RP010.Click, RP009.Click, RP008.Click, RP007.Click, RP006.Click, RP005.Click, RP004.Click, RP003.Click, RP002.Click, RP001.Click
        {
            assignDefaultButton();

            Button btnLoc = sender as Button;
            String kb_loc = btnLoc.Name;
            String site = "AT";
            btnLoc.BackColor = Color.RoyalBlue;
            DataRow[] dr = this.kanbanLayoutSecition1.dsKanbanLayout.Tables[0].Select("kb_loc='" + kb_loc + "'");
            String part_num = "";
            int bin_qty = 0;
            if (dr.Length > 0)
            {
                 part_num = dr[0]["partnum"].ToString();
                 bin_qty = Convert.ToInt32(dr[0]["binqty"].ToString());
            }
            KB_PopupOrderAssign kb = new KB_PopupOrderAssign(kb_loc, site, part_num, bin_qty);
            kb.ShowDialog();
            kanbanLayoutSecition1.getdataSet();
            btnLoc.BackColor = Color.RoyalBlue;
          
        }
 
    }
}
