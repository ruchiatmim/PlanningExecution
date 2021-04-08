using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Deployment.Application;
using System.Configuration;

namespace Version3
{
    public partial class PartSearch : Form
    {

        DataTable tbl = new DataTable();
        DataTable tblFinal = new DataTable();
        DataSet ds = new DataSet();
        string constr = ConfigurationManager.ConnectionStrings["MVAPPSConnectionString"].ConnectionString;
        public PartSearch()
        {
            InitializeComponent();
        }

        private void RegexMapping_Load(System.Object sender, System.EventArgs e)
        {

           // statusname.Text = "";
          //  statusname.Text = My.Settings.UserID;
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                ApplicationDeployment deploy = ApplicationDeployment.CurrentDeployment;
              //  statusversion.Text = deploy.CurrentVersion.ToString();
            }
            txtManfPartNum.Select();
        }
        private void SearchRegex()
{
        Regex reg ;
        try {
            tblFinal.Rows.Clear();
        }
        catch(Exception e)
      {}
        try {
            tbl.Rows.Clear();
        }
        catch (Exception e)
         {}

        try {
            tblFinal.Columns.Add("ID", System.Type.GetType("System.String"));
            tblFinal.Columns.Add("Part Number", System.Type.GetType("System.String"));
            tblFinal.Columns.Add("R_Part Number", Type.GetType("System.String"));
            tblFinal.Columns.Add("Category", Type.GetType("System.String"));
            tblFinal.Columns.Add("Operation", Type.GetType("System.String"));
            tblFinal.Columns.Add("MRB", Type.GetType("System.String"));
        }
       catch
        {}

        //'open new connection with connection string
        SqlConnection conn = new System.Data.SqlClient.SqlConnection(constr);
        conn.Open();

        //'Dim strVal As String = strSql.ToString()
        String strVal  = "exec ControlDb.dbo.sf_getMappingPartDetails";
        SqlCommand cmd = new SqlCommand(strVal, conn);
        SqlDataReader rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
        tbl.Load(rdr);

        if (tbl.Rows.Count > 0 )
        {
            foreach (DataRow dr   in tbl.Rows)
            {
                reg = new Regex(dr["manf_part_reg"].ToString());
                if (reg.IsMatch(txtManfPartNum.Text.Trim()))
                {
                    if (dr["mrb_status"].ToString().Equals("Y"))
                    {
                       DataRow drtable=tblFinal.NewRow();
                        drtable["ID"]=dr["ID"].ToString();
                        drtable["Part Number"]=dr["part_num"].ToString();
                        drtable["R_Part Number"]=dr["r_part_num"].ToString();
                        drtable["Category"]=dr["category_desc"].ToString();
                        drtable["Operation"]=dr["oper_desc"].ToString();
                        drtable["MRB"]="Yes";
                         tblFinal.Rows.Add(drtable);
                       // tblFinal.Rows.Add(New Object() {dr("ID").ToString(), dr("part_num").ToString(), dr("r_part_num").ToString(), dr("category_desc").ToString(), dr("oper_desc").ToString(), "Yes"})
                    }
                     else 
                        {
                        DataRow drtable=tblFinal.NewRow();
                        drtable["ID"]=dr["ID"].ToString();
                        drtable["Part Number"]=dr["part_num"].ToString();
                        drtable["R_Part Number"]=dr["r_part_num"].ToString();
                        drtable["Category"]=dr["category_desc"].ToString();
                        drtable["Operation"]=dr["oper_desc"].ToString();
                        drtable["MRB"]="No";
                       tblFinal.Rows.Add(drtable);
                       // tblFinal.Rows.Add(new Object() {dr("ID").ToString(), dr("part_num").ToString(), dr("r_part_num").ToString(), dr("category_desc").ToString(), dr("oper_desc").ToString(), "No"})
                    }
                    lblDescription.Text = dr["description"].ToString();
                    lblDescription.ForeColor = Color.SaddleBrown;
                    lblDescription.AutoSize = true;
                }
         }

            if (tblFinal.Rows.Count > 0 )
            {
                gridview.ColumnHeadersVisible = true;
                gridview.DataSource = tblFinal;
                gridview.AllowUserToAddRows = false;
                gridview.AllowUserToDeleteRows = false;
                gridview.Columns["ID"].Width = 50;
                gridview.Columns["ID"].ReadOnly = true;

                gridview.Columns["Part Number"].Width = 150;
                gridview.Columns["Part Number"].ReadOnly = true;
                gridview.Columns["Part Number"].HeaderText = "Normal Part Number";
                gridview.Columns["R_Part Number"].Width = 150;
                gridview.Columns["R_Part Number"].ReadOnly = true;
                gridview.Columns["R_Part Number"].HeaderText = "Used Part Number";
                gridview.Columns["Category"].Width = 145;
                gridview.Columns["Category"].ReadOnly = true;
                gridview.Columns["Category"].HeaderText = "Build Operation";
                gridview.Columns["Operation"].Width = 145;
                gridview.Columns["Operation"].ReadOnly = true;
                gridview.Columns["Operation"].HeaderText = "Decom Operation";
                gridview.Columns["MRB"].ReadOnly = true;

                gridview.CurrentRow.Selected = false;
                gridview.CurrentCell.Selected = false;
            }
            else
                {
                lblScannedPart.Text = "Part Number Not Found";
                lblScannedPart.ForeColor = Color.SaddleBrown;
                lblDescription.Text = "";
                }
                try {
                    gridview.ColumnHeadersVisible = false;
                }
                catch (Exception ex)
                    {}
                
            }
        }
        //    txtManfPartNum.Clear();
        //    txtManfPartNum.Select();


        private void btnSearch_Click(System.Object sender, System.EventArgs e)
        {

            if (!txtManfPartNum.Text.Trim().Equals(""))
            {
                SearchRegex();
            }
            else
            {
                //'try {
                //'    tblFinal.Rows.Clear()
                //'Catch
                //'} {
                //'try {
                //'    tbl.Rows.Clear()
                //'Catch
            }
            try
            {
                gridview.DataSource = tblFinal;
            }
            catch (Exception ex)
            { }
        }





        private void txtManfPartNum_LostFocus(Object sender, System.EventArgs e)
        {
            if (!txtManfPartNum.Text.Trim().Equals(""))
            {
                lblDescription.Text = "";
                lblScannedPart.Text = txtManfPartNum.Text.Trim();
                lblScannedPart.ForeColor = Color.SaddleBrown;
                SearchRegex();
            }
            else
            {
                //'try {
                //'    tblFinal.Rows.Clear()
                //'Catch
                //'} {
                //'try {
                //'    tbl.Rows.Clear()
                //'Catch
                //'} {
                //'try {
                //'    gridview.DataSource = tblFinal
                //'Catch ex As Exception
                //'} {
            }
        }


        private void gridview_Click(System.Object sender, System.EventArgs e)
        {
            gridview.CurrentCell.Selected = false;
            gridview.CurrentRow.Selected = false;
        }

        private void gridview_MouseClick(System.Object sender, System.Windows.Forms.MouseEventArgs e)
        {
            gridview.CurrentCell.Selected = false;
            gridview.CurrentRow.Selected = false;

        }

        private void gridview_CellMouseClick(System.Object sender, System.Windows.Forms.DataGridViewCellMouseEventArgs e)
        {
            gridview.CurrentCell.Selected = false;
            gridview.CurrentRow.Selected = false;

        }

        private void gridview_CellMouseDown(System.Object sender, System.Windows.Forms.DataGridViewCellMouseEventArgs e)
        {
            gridview.CurrentCell.Selected = false;
            gridview.CurrentRow.Selected = false;
        }

        private void gridview_CellStateChanged(System.Object sender, System.Windows.Forms.DataGridViewCellStateChangedEventArgs e) 
{
        if (e.StateChanged== DataGridViewElementStates.Selected)
        {
            e.Cell.Selected = false;
        }

    }

        private void gridview_RowPrePaint(System.Object sender, System.Windows.Forms.DataGridViewRowPrePaintEventArgs e)
        {
            // ' e.PaintParts &= DataGridViewPaintParts.Focus
            e.PaintParts &= ~DataGridViewPaintParts.Focus;


        }

      
    }


}