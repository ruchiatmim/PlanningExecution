using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Data.SqlClient;
using System.Configuration;
namespace Version3
{
    public partial class KanbanLayoutSecition : UserControl
    {
        public KanbanLayoutSecition()
        {
            InitializeComponent();
            getdataSet();
        }
        string site;
        string constr = "Data Source=mverp;Initial Catalog=MIMDIST;user id=sa;password=mimi~100;";
      
        public DataSet dsKanbanLayout;
        public void getdataSet()
         {
        site = "AT";
        SqlConnection conn = new System.Data.SqlClient.SqlConnection(constr);
        conn.Open();
        String cmdtxt = "SELECT  [PartNum] ,[KB_Loc] ,binQTY,contType,sector       FROM  [dbo].[v_KBLayout_PartAssignment] where [SiteCode] ='" + site + "' order by PartNum";

        SqlCommand cmd  = new SqlCommand(cmdtxt, conn);
        cmd.CommandTimeout = 0;
        SqlDataAdapter da  = new SqlDataAdapter(cmd);
        dsKanbanLayout = new DataSet();
        da.Fill(dsKanbanLayout);      
        assignDefaultButton();
         }


        DataTable dtLayout;
        int xloc, yloc;
        private void assignDefaultButton()
        {
     
      
      

        try
        {
            foreach (DataRow dr in dsKanbanLayout.Tables[0].Rows)
            {
                String kb_loc;
                String PartNum;
                kb_loc = dr["KB_Loc"].ToString();
                PartNum = dr["PartNum"].ToString();
                Button btnloc = new Button();
                if (PanelA.Controls.Find(kb_loc, false).Length > 0)
                    btnloc = PanelA.Controls.Find(kb_loc, false)[0] as Button;
                else if (PanelB.Controls.Find(kb_loc, false).Length > 0)
                    btnloc = PanelB.Controls.Find(kb_loc, false)[0] as Button;
                else if (PanelC.Controls.Find(kb_loc, false).Length > 0)
                    btnloc = PanelC.Controls.Find(kb_loc, false)[0] as Button;
                else if (PanelD.Controls.Find(kb_loc, false).Length > 0)
                    btnloc = PanelD.Controls.Find(kb_loc, false)[0] as Button;
                else
                {
                    Label LBL = new Label();
                    LBL.Text = kb_loc;
                    LBL.Location = new System.Drawing.Point(xloc, yloc);
                    PanelNF.Controls.Add(LBL);
                    yloc = yloc + 30;

                    PanelNF.Visible = true;
                    lblNF.Visible = true;

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
   
       catch (Exception ex )
            {

            //'If btnloc Is Nothing Then
            //'    Dim LBL As New Label
            //'    LBL.Text = kb_loc
            //'    PanelNF.Controls.Add(LBL)

            //'End If

        }

        }


     


      

        private void KanbanLayoutSecition_Paint(object sender, PaintEventArgs e)
        {
              //'Secion A
             dtLayout = new DataTable();
            dtLayout.Columns.Add(new DataColumn("Bin"));
            dtLayout.Columns.Add(new DataColumn("Section"));

            DataRow dr = dtLayout.NewRow();
            dr["Bin"] = "BR066";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BR065";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

             dr = dtLayout.NewRow();
             dr["Bin"] = "BR064";
             dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BR063";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

           dr = dtLayout.NewRow();
           dr["Bin"] = "BR062";
           dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "BR061";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "BP060";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "BP059";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP058";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP057";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP056";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP055";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP054";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP053";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP052";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP051";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP050";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP049";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP048";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP047";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP046";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP045";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP044";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP043";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP042";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP041";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);
            dr = dtLayout.NewRow();
            dr["Bin"] = "BP040";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP039";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP038";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);
           
          
            dr = dtLayout.NewRow();
            dr["Bin"] = "BP037";
            dr["Section"] = "A1";
            dtLayout.Rows.Add(dr);



            // Top face
                   

            dr = dtLayout.NewRow();
            dr["Bin"] = "BR005";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BR006";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BR007";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BR008";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BR009";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BR010";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP011";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP012";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "BP013";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP014";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP015";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);
           

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP016";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP017";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP018";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP019";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP020";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP021";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP022";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP023";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP024";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP025";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP026";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP027";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);
         
            dr = dtLayout.NewRow();
            dr["Bin"] = "BP028";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP029";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP030";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "BP031";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP032";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);
         
            dr = dtLayout.NewRow();
            dr["Bin"] = "BP033";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP034";
            dr["Section"] = "A";
            dtLayout.Rows.Add(dr);

             
           

            //***********SECTION B********************************//
            dr = dtLayout.NewRow();
            dr["Bin"] = "BP071";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);
       

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP072";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP073";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP074";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP075";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP076";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP077";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP078";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP079";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP080";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "BP081";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP082";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP083";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP084";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP085";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP086";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP087";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP088";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "BP089";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP090";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP091";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP092";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP093";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP094";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP095";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP096";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP097";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP098";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP099";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "BP100";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP101";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP102";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP103";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP104";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP105";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "BP106";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP107";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP108";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP109";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "BP110";
            dr["Section"] = "B";
            dtLayout.Rows.Add(dr);



            //***********SECTION D******************************//
            dr = dtLayout.NewRow();
            dr["Bin"] = "RR001";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "RR002";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RR003";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RR004";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "RR005";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RR006";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RR007";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "RR008";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RR009";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RR010";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "RR011";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "RR012";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RR013";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RR014";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "RR015";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RR016";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RR017";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "RR018";
            dr["Section"] = "D";
            dtLayout.Rows.Add(dr);




            //***********SECTION C********************************//
            dr = dtLayout.NewRow();
            dr["Bin"] = "RP001";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP002";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP003";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP004";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP005";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP006";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP007";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP008";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP009";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP010";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP011";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP012";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP013";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP014";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP015";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP016";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP017";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP018";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP019";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP020";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP021";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "RP022";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP023";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP024";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "RP025";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP026";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP027";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);


            dr = dtLayout.NewRow();
            dr["Bin"] = "RP028";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);

            dr = dtLayout.NewRow();
            dr["Bin"] = "RP029";
            dr["Section"] = "C";
            dtLayout.Rows.Add(dr);



        e.Graphics.TranslateTransform(158, 615);
      //  Font f  = new Font("Verdana", 12,);
        Font f = new System.Drawing.Font("Verdana", 14F, System.Drawing.FontStyle.Regular);

        DataRow[] drSecA1 = dtLayout.Select("SECTION='A1'"); 
        DataRow[] drSecA=dtLayout.Select("SECTION='A'");
        DataRow[] drSecB = dtLayout.Select("SECTION='B'");
        DataRow[] drSecC = dtLayout.Select("SECTION='C'");
        DataRow[] drSecD = dtLayout.Select("SECTION='D'");


        //'*******************************************************************************************
        //'*************************************SECTION A**** BOTTOM LABEL****************************
        //'*******************************************************************************************
    int secA_gap1=35;
    int secA_gap2=44;
    int ind = 0;
    e.Graphics.RotateTransform(270);
    foreach (DataRow drA in drSecA1)
    {
        ind += 1;
      
        string lbl = drA[0].ToString();
        e.Graphics.DrawString(lbl, f, Brushes.Black, this.ClientRectangle);
        if (ind % 2 == 0)
            e.Graphics.TranslateTransform(0, secA_gap2);
        else
            e.Graphics.TranslateTransform(0, secA_gap1);
       
    }

  //  e.Graphics.TranslateTransform(275, -1143);

    e.Graphics.TranslateTransform(275, -1183);
    foreach (DataRow drA in drSecA)
    {
        ind += 1;
        string lbl = drA[0].ToString();
        e.Graphics.DrawString(lbl, f, Brushes.Black, this.ClientRectangle);
        if (ind % 2 == 0)
            e.Graphics.TranslateTransform(0, secA_gap2);
        else
            e.Graphics.TranslateTransform(0, secA_gap1);

    }

    int secB_gap1 = 30;
    int secB_gap2 = 44;
    int secB_beamgap = 36;
    e.Graphics.TranslateTransform(-520, -1225);
  // e.Graphics.TranslateTransform(275, -1183);
    ind = 0;
    foreach (DataRow drA in drSecB)
    {
        ind += 1;
        string lbl = drA[0].ToString();
        e.Graphics.DrawString(lbl, f, Brushes.Black, this.ClientRectangle);
        if (lbl=="BP080" || lbl=="BP090" || lbl=="BP100")
            e.Graphics.TranslateTransform(0, secB_beamgap);
        else
            e.Graphics.TranslateTransform(0, secB_gap1);
     

    }

    e.Graphics.TranslateTransform(655, -780);
    foreach (DataRow drA in drSecD)
    {
        ind += 1;
        string lbl = drA[0].ToString();
        e.Graphics.DrawString(lbl, f, Brushes.Black, this.ClientRectangle);
      
            e.Graphics.TranslateTransform(0, secA_gap1);

    }

    //e.Graphics.RotateTransform(270);



    //e.Graphics.DrawString("BR066", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);


    //e.Graphics.DrawString("BR065", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BR064", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BR063", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BR062", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BR061", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP060", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP059", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP058", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP057", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP056", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP055", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP054", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP053", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP052", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP051", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP050", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP049", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP048", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP047", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP046", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP045", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP044", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP043", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP042", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP041", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP040", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP039", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap2);
    //e.Graphics.DrawString("BP038", f, Brushes.Black, this.ClientRectangle);
    //e.Graphics.TranslateTransform(0, secA_gap1);
    //e.Graphics.DrawString("BP037", f, Brushes.Black, this.ClientRectangle);

  //  SectionB:

   //     '*******************************************************************************************
  //      '*************************************SECTION A**** TOP LABEL*******************************
  //      '*******************************************************************************************
  //
            //e.Graphics.TranslateTransform(275, -1143);
            //e.Graphics.DrawString("BR005", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BR006", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BR007", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BR008", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BR009", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BR010", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP011", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP012", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP013", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP014", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP015", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP016", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP017", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP018", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP019", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP020", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP021", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP022", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP023", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP024", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP025", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP026", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP027", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP028", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP029", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP030", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP031", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP032", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap2);
            //e.Graphics.DrawString("BP033", f, Brushes.Black, this.ClientRectangle);
            //e.Graphics.TranslateTransform(0, secA_gap1);
            //e.Graphics.DrawString("BP034", f, Brushes.Black, this.ClientRectangle);
       

        //       '*******************************************************************************************
        //       '*************************************SECTION B ********************************************
        //      '*******************************************************************************************


        //int secB_gap1=30;
        //int secB_gap2=44;
        //int secB_beamgap=36;
        //e.Graphics.TranslateTransform(-520, -1175);
        ////     ' e.Graphics.RotateTransform(270)
        //e.Graphics.DrawString("BP071", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, 31);
        //e.Graphics.DrawString("BP072", f, Brushes.Black, this.ClientRectangle);
    
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP073", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP074", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP075", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP076", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP077", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP078", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP079", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP080", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_beamgap);
        //e.Graphics.DrawString("BP081", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP082", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP083", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP084", f, Brushes.Black, this.ClientRectangle);

        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP085", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP086", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP087", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP088", f, Brushes.Black, this.ClientRectangle);

        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP089", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP090", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_beamgap);
        //e.Graphics.DrawString("BP091", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP092", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP093", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP094", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP095", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP096", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP097", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP098", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP099", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP100", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_beamgap);
        //e.Graphics.DrawString("BP101", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP102", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP103", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP104", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP105", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP106", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP107", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP108", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP109", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secB_gap1);
        //e.Graphics.DrawString("BP110", f, Brushes.Black, this.ClientRectangle);
       

        //   '*******************************************************************************************
        //    '*************************************SECTION D LABEL*********************************************
        //    '*******************************************************************************************
        //e.Graphics.TranslateTransform(655, -755);


        ////  ' e.Graphics.RotateTransform(270)
        //e.Graphics.DrawString("RR001", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR002", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR003", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR004", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR005", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR006", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR007", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR008", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR009", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR010", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR011", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR012", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR013", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR014", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR015", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR016", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR017", f, Brushes.Black, this.ClientRectangle);
        //e.Graphics.TranslateTransform(0, secA_gap1);
        //e.Graphics.DrawString("RR018", f, Brushes.Black, this.ClientRectangle);
    xloc = 10;
    yloc = 10;
    //PanelNF.Visible = false;
    //lblNF.Visible = false;

        foreach (DataRow drlay in dtLayout.Rows)
        {
            if (dsKanbanLayout.Tables[0].Select("KB_Loc='" + drlay[0].ToString()+"'").Length < 1)
            {
                Label LBL = new Label();
                LBL.Text = drlay[0].ToString();
                LBL.Location = new System.Drawing.Point(xloc, yloc);
                PanelNF.Controls.Add(LBL);
                yloc = yloc + 30;

                PanelNF.Visible = true;
                lblNF.Visible = true;
            }
        }


        foreach (DataRow drlay in dsKanbanLayout.Tables[0].Rows)
        {
            if (dtLayout.Select("bin='" + drlay["KB_Loc"].ToString() + "'").Length < 1)
            {
                Label LBL = new Label();
                LBL.Text = drlay["KB_Loc"].ToString();
                LBL.Location = new System.Drawing.Point(xloc, yloc);
                PanelNF.Controls.Add(LBL);
                yloc = yloc + 30;

                PanelNF.Visible = true;
                lblNF.Visible = true;
            }
        }
        }
  
    
    }
}
