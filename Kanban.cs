using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Deployment.Application;
using System.Configuration;
using System.Threading;


using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading.Tasks;
using Google.Apis.Download;

using System.Diagnostics;


namespace Version3
{
    public partial class Kanban : Form
    {
        string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
        public Kanban()
        {
            InitializeComponent();
        }

        public DataSet getdataSet(string sql)
        {
            SqlConnection cn = new SqlConnection(constr);
            SqlCommand cmd = new SqlCommand(sql, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet dsSerial = new DataSet();
            da.Fill(dsSerial);
            return dsSerial;
        }

        private void Kanban_Load(object sender, EventArgs e)
        {
            loadFilterSiteCode();
            loadSiteCode();
            loadOrderPrefix();
            loadOrderType();
            loadstatus();
            loadContainerType();
            loadPartType();
            // loadGrid();         
        }

        public void loadPartType()
        {
            string sql1 = " SELECT  distinct TypeCode,part_type_description  FROM    m_app_kanban_master    ";
            DataTable dt = getdataSet(sql1).Tables[0];
            comboPartType.DataSource = dt;
            comboPartType.DisplayMember = "part_type_description";
            comboPartType.ValueMember = "TypeCode";
            comboPartType.SetDataBinding(dt, null);
            comboPartType.DisplayLayout.Bands[0].Columns["TypeCode"].Hidden = true;
            comboPartType.DisplayLayout.Bands[0].ColHeadersVisible = false;

            string sqlRep = " SELECT  distinct TypeCode,part_type_description  FROM    m_app_Replen_master   ";
            DataTable dtReplenish = getdataSet(sqlRep).Tables[0];
            comboPartTypeRep.DataSource = dtReplenish;
            comboPartTypeRep.DisplayMember = "part_type_description";
            comboPartTypeRep.ValueMember = "TypeCode";
            comboPartTypeRep.SetDataBinding(dtReplenish, null);
            comboPartTypeRep.DisplayLayout.Bands[0].Columns["TypeCode"].Hidden = true;
            comboPartTypeRep.DisplayLayout.Bands[0].ColHeadersVisible = false;

            string sqldyn = " SELECT  distinct TypeCode,part_type_description  FROM    m_app_dynamic_master   ";
            DataTable dtDynamic = getdataSet(sqldyn).Tables[0];
            this.ddFilterdynPartType.DataSource = dtDynamic;
            ddFilterdynPartType.DisplayMember = "part_type_description";
            ddFilterdynPartType.ValueMember = "TypeCode";
            ddFilterdynPartType.SetDataBinding(dtDynamic, null);
            ddFilterdynPartType.DisplayLayout.Bands[0].Columns["TypeCode"].Hidden = true;
            ddFilterdynPartType.DisplayLayout.Bands[0].ColHeadersVisible = false;
        }


        public void loadGrid()
        {
            string sql = " SELECT  *  FROM    m_app_kanban_master where TOrdType='KANBAN'  ";
            if (ddFilterSite.Value != null)
                if (ddFilterSite.Value.ToString() != "ALL")
                    sql = sql + "  and  TSiteCode='" + ddFilterSite.SelectedText.ToString() + "'";
            if (this.comboPartType.Value != null)
                if (comboPartType.Value.ToString() != "ALL" && comboPartType.Value.ToString() != "")
                    sql = sql + "  and  TypeCode='" + comboPartType.Value.ToString() + "'";
            if (this.txtPartNum.Text.Trim() != "")
                sql = sql + "  and  TPartNum like '" + this.txtPartNum.Text.Trim() + "%'";
            if (this.ultraTextEditorDescription.Text.Trim() != "")
                sql = sql + "  and  part_description like '%" + this.ultraTextEditorDescription.Text.Trim() + "%'";

            sql = sql + " order by TPartNum ";
            DataTable dt = getdataSet(sql).Tables[0];
            this.gridKanban.DataSource = dt;
            lblNumRecKan.Text = " Num of Rec: " + dt.Rows.Count.ToString();
            
            //  bindingSource1.DataSource = getdataSet(sql).Tables[0];
            //  bindingNavigator1.BindingSource = bindingSource1;

        }


        public void loadReplenishGrid()
        {
            string sql = "SELECT  *  FROM  m_app_Replen_master  where TPartNum !=''    ";

            if (ddFilterRepSite.Value != null)
                if (ddFilterRepSite.Value.ToString() != "ALL")
                    sql = sql + "  and  TSiteCode='" + this.ddFilterRepSite.SelectedText.ToString() + "'";

            if (this.comboPartTypeRep.Value != null)
                if (comboPartTypeRep.Value.ToString() != "ALL" && comboPartTypeRep.Value.ToString() != "")
                    sql = sql + "  and  TypeCode='" + comboPartTypeRep.Value.ToString() + "'";

            if (this.txtRepPartNum.Text.Trim() != "")
                sql = sql + "  and  TPartNum like '" + this.txtRepPartNum.Text.Trim() + "%'";

            if (this.comboOrderType.Text.Trim() != "")
                sql = sql + "  and   TOrdType ='" + this.comboOrderType.Text.Trim() + "'";

            if (this.ultraTextEditorRepDescription.Text.Trim() != "")
                sql = sql + "  and  part_description like '%" + this.ultraTextEditorRepDescription.Text.Trim() + "%'";


            sql = sql + " order by TPartNum ";
            DataTable dt = getdataSet(sql).Tables[0];
            this.gridReplenish.DataSource = dt;
            lblNumRec.Text = " Num of Rec: " + dt.Rows.Count.ToString();
            //  bindingSource1.DataSource = getdataSet(sql).Tables[0];
            //  bindingNavigator1.BindingSource = bindingSource1;
        }


        public void loadDynamicGrid()
        {
            string sql = "SELECT  *  FROM  m_app_dynamic_master  where TPartNum !=''    ";

            if (ddFilterdynSite.Value != null)
                if (ddFilterdynSite.Value.ToString() != "ALL")
                    sql = sql + "  and  TSiteCode='" + this.ddFilterdynSite.SelectedText.ToString() + "'";

            if (this.ddFilterdynPartType.Value != null)
                if (ddFilterdynPartType.Value.ToString() != "ALL" && ddFilterdynPartType.Value.ToString() != "")
                    sql = sql + "  and  TypeCode='" + ddFilterdynPartType.Value.ToString() + "'";

            if (this.txtFilterdynPartNum.Text.Trim() != "")
                sql = sql + "  and  TPartNum like '" + this.txtFilterdynPartNum.Text.Trim() + "%'";

            if (this.ultraTextEditorDynDescription.Text.Trim() != "")
                sql = sql + "  and  part_description like '%" + this.ultraTextEditorDynDescription.Text.Trim() + "%'";

            //   if (this.ddFilterOrdType.Text.Trim() != "")
            //      sql = sql + "  and   TOrdType ='" + this.ddFilterOrdType.Text.Trim() + "'";

            sql = sql + " order by TPartNum ";
            DataTable dt = getdataSet(sql).Tables[0];
            this.grdDynamic.DataSource = dt;
            lblNumRec.Text = " Num of Rec: " + dt.Rows.Count.ToString();
            //  bindingSource1.DataSource = getdataSet(sql).Tables[0];
            //  bindingNavigator1.BindingSource = bindingSource1;
        }


        private void gridKanban_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
        {
            /*     if (e.Row.Cells["NEW"].Value.ToString() == "false")
             //    e.Row.Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
             {
                 //   this.gridKanban.Rows[e.Row.Index].Cells[0].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                 this.gridKanban.Rows[e.Row.Index].Cells[1].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                 this.gridKanban.Rows[e.Row.Index].Cells[2].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                 this.gridKanban.Rows[e.Row.Index].Cells[3].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                 this.gridKanban.Rows[e.Row.Index].Cells[4].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                 this.gridKanban.Rows[e.Row.Index].Cells[5].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                 this.gridKanban.Rows[e.Row.Index].Cells[6].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                 this.gridKanban.Rows[e.Row.Index].Cells[7].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                 this.gridKanban.Rows[e.Row.Index].Cells[8].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                 this.gridKanban.Rows[e.Row.Index].Cells[9].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                 this.gridKanban.Rows[e.Row.Index].Cells["TDelete"].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                 this.gridKanban.Rows[e.Row.Index].Cells["Save"].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
             }
             else
             {
                 this.gridKanban.Rows[e.Row.Index].Cells["NEW"].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
             }
           //  if (e.Row.Cells["TDelete"].Value == "Y")
             this.gridKanban.Rows[e.Row.Index].Cells["Save"].Value = "Save";
           */

        }

        public void loadFilterSiteCode()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Site Code"));
            DataRow dr = dt.NewRow();
         /*   dr[0] = "ALL";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr[0] = "AT";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr[0] = "FB";
            dt.Rows.Add(dr);*/
          //  dr = dt.NewRow();
            dr[0] = "NL";
            dt.Rows.Add(dr);
          //  dr = dt.NewRow();
          //  dr[0] = "SG";
          //  dt.Rows.Add(dr);
            this.ddFilterSite.DataSource = dt;
            ddFilterRepSite.DataSource = dt;
            this.ddFilterdynSite.DataSource = dt;
        }


        public void loadSiteCode()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Site Code"));
            DataRow dr = dt.NewRow();
     
        dr[0] = "NL";
            dt.Rows.Add(dr);
        //    dr = dt.NewRow();
        //    dr[0] = "SG";
        //    dt.Rows.Add(dr);
            this.ddSite.DataSource = dt;
            comboSite.DataSource = dt;
            comboRepSite.DataSource = dt;
            //ddFilterdynSite.DataSource = dt;
        }

        public void loadstatus()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Status"));
            DataRow dr = dt.NewRow();
            dr[0] = "N";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr[0] = "R";
            dt.Rows.Add(dr);
            this.ddStatus.DataSource = dt;
        }

        public void loadOrderType()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Ord Type"));
            DataRow dr = dt.NewRow();
            dr[0] = "KANBAN";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr[0] = "REPLENISH";
            dt.Rows.Add(dr);
            this.ddOrderType.DataSource = dt;
            comboOrderType.DataSource = dt;
          //  this.ddFilterOrdType.DataSource = dt;

        }

        public void loadOrderPrefix()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Ord Prefix"));
            DataRow dr = dt.NewRow();
            dr[0] = "XBKNBA";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr[0] = "XBKNBC";
            dt.Rows.Add(dr);
            this.ddOrderPrefix.DataSource = dt;
         //   this.comboPrefix.DataSource = dt;
            // this.comboRepOrdPrefix.DataSource = dt;

            string sqlCont = "select distinct conttype from KanBanLayoutMaster";


            DataTable dtCont = getdataSet(sqlCont).Tables[0];

            int xloc, yloc;
            xloc = 10;
            yloc = 5;

            foreach (DataRow drPref in dt.Rows)
            {
                RadioButton rad = new RadioButton();
                rad.Text = drPref["Ord Prefix"].ToString();
                rad.Name = drPref["Ord Prefix"].ToString();
                rad.Location = new Point(xloc, yloc);
                xloc = xloc + 120;
                pnlOrderPrefix.Controls.Add(rad);
            }


        }

        public void loadContainerType()
        {
          //  string sqlCont="select distinct conttype from KanBanLayoutMaster";


         //   DataTable dtCont = getdataSet(sqlCont).Tables[0];

            DataTable dtCont = new DataTable();
            dtCont.Columns.Add(new DataColumn("contType"));

            DataRow dr = dtCont.NewRow();
                  
            dr[0] = "BATTERY";
            dtCont.Rows.Add(dr);
            
            dr = dtCont.NewRow();
            dr[0] = "BOX";
            dtCont.Rows.Add(dr);

            dr = dtCont.NewRow();
            dr[0] = "BULK";
            dtCont.Rows.Add(dr);

            dr = dtCont.NewRow();
            dr[0] = "LRGRACK";
            dtCont.Rows.Add(dr);
            
            dr = dtCont.NewRow();
            dr[0] = "PALLET";
            dtCont.Rows.Add(dr);                 
            
            dr = dtCont.NewRow();
            dr[0] = "STDRACK";
            dtCont.Rows.Add(dr);
  
            int xloc,yloc;
            xloc=10;
            yloc=0;

            foreach (DataRow drCont in dtCont.Rows)
            {
                RadioButton rad=new RadioButton();
                rad.Text = drCont["contType"].ToString();
                rad.Name = drCont["contType"].ToString();
                rad.Location = new Point(xloc, yloc);
                yloc=yloc+20;
                pnlContiner.Controls.Add(rad);
            }

            //ddContainerType.DataSource = dt;
            //comboContType.DataSource = dt;
            //comboContType.SelectedText = "SELECT";
            //   this.comboRepContType.DataSource = dt;
        }

        private void gridKanban_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
        {
            /*  if (e.Cell.Column.Index == 12)
              {
                  string status = "";
                   int id = Convert.ToInt32(e.Cell.Row.Cells["MasterID"].Value.ToString());
                   if (id > 0)
                   {
                       string FlagDelete = e.Cell.Row.Cells["TDelete"].Text.ToString();
                       if (FlagDelete == "True")
                       {
                           status = "Y";
                           //  else
                           //    status = "N";
                           DeleteRow(id, status);
                       }
                       loadGrid();
                   }

              }*/
            if ((e.Cell.Column.Index == 6) || (e.Cell.Column.Index == 7))
            {
                int QtyInCont = 0;
                int QtyOfOrders = 0;
                if (e.Cell.Row.Cells["TQtyInCont"].Text.Trim().ToString() != "")
                {
                    QtyInCont = Convert.ToInt32(e.Cell.Row.Cells["TQtyInCont"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                }
                if (e.Cell.Row.Cells["TQtyOfOrders"].Text.Trim().ToString() != "")
                {
                    QtyOfOrders = Convert.ToInt32(e.Cell.Row.Cells["TQtyOfOrders"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                }
                int tot_qty = QtyInCont * QtyOfOrders;

                e.Cell.Row.Cells["Tot_Qty"].Value = tot_qty;
            }
            /*   else if (e.Cell.Column.Index == 10)
               {
                
                   if (e.Cell.Row.Cells["MasterID"].Value.ToString() != "")
                   {
                       int id = Convert.ToInt32(e.Cell.Row.Cells["MasterID"].Value.ToString());
                       if (id > 0)
                       {                        
                           string FlagNew = e.Cell.Row.Cells["NEW"].Text.ToString();
                           //   CheckBox chk=(CheckBox)(e.Cell.Row.Cells["NEW"]);
                           string status = "";// e.Cell.Row.Cells["Tstatus"].Value.ToString();
                           if (FlagNew.ToUpper() == "TRUE")
                           {
                               status = "N";
                               //  else
                               //     status = "R";
                               string site_code = e.Cell.Row.Cells["TSiteCode"].Text.ToString();
                               string part_num = e.Cell.Row.Cells["TPartNum"].Text.ToString();
                               string displayPartNum = e.Cell.Row.Cells["TDisplayPN"].Text.ToString();
                               string OrdType = e.Cell.Row.Cells["TOrdType"].Text.ToString();
                               string OrdPrefix = e.Cell.Row.Cells["TOrdPrefix"].Text.ToString();
                               string ContType = e.Cell.Row.Cells["TContType"].Text.ToString();
                               int QtyInCont = Convert.ToInt32(e.Cell.Row.Cells["TQtyInCont"].Text.ToString());
                               int QtyOfOrders = Convert.ToInt32(e.Cell.Row.Cells["TQtyOfOrders"].Text.ToString());
                               UpdateRow(id, site_code, part_num, displayPartNum, OrdType, OrdPrefix, ContType, QtyInCont, QtyOfOrders, status);
                               loadGrid();
                               e.Cell.Row.Cells["NEW"].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
                           }
                       }
                   }
               }*/
        }
        public void UpdateRow(int id, string site_code, string part_num, string displayPartNum, string OrdType, string OrdPrefix, string ContType, int QtyInCont, int QtyOfOrders, string status)
        {
            string sqlUpdate = "Update KanbanMasterTemp set TSiteCode='" + site_code + "',TPartNum='" + part_num + "',TDisplayPN='" + displayPartNum + "', TOrdType='" + OrdType + "', TOrdPrefix='" + OrdPrefix + "', TContType='" + ContType + "', TQtyInCont='" + QtyInCont + "', TQtyOfOrders='" + QtyOfOrders + "', Tstatus='" + status + "'  where MasterID=" + id;
            getdataSet(sqlUpdate);

        }


        private void ugData_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e)
        {
            e.Cancel = true;
        }


        private void gridKanban_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
        {
            /*  if (e.Cell.Row.Cells["MasterID"].Value.ToString() != "")
                {
                    int id = Convert.ToInt32(e.Cell.Row.Cells["MasterID"].Value.ToString());
                    if (id > 0)
                    {
                        string FlagNew = e.Cell.Row.Cells["NEW"].Text.ToString();
                        //   CheckBox chk=(CheckBox)(e.Cell.Row.Cells["NEW"]);
                        string status = "";// e.Cell.Row.Cells["Tstatus"].Value.ToString();
                        if (FlagNew.ToUpper() == "TRUE")
                            status = "N";
                        else
                            status = "R";
                        string site_code = e.Cell.Row.Cells["TSiteCode"].Text.ToString();
                        string part_num = e.Cell.Row.Cells["TPartNum"].Text.ToString();
                        string displayPartNum = e.Cell.Row.Cells["TDisplayPN"].Text.ToString();
                        string OrdType = e.Cell.Row.Cells["TOrdType"].Text.ToString();
                        string OrdPrefix = e.Cell.Row.Cells["TOrdPrefix"].Text.ToString();
                        string ContType = e.Cell.Row.Cells["TContType"].Text.ToString();
                        int QtyInCont = Convert.ToInt32(e.Cell.Row.Cells["TQtyInCont"].Text.ToString());
                        int QtyOfOrders = Convert.ToInt32(e.Cell.Row.Cells["TQtyOfOrders"].Text.ToString());
                        UpdateRow(id, site_code, part_num, displayPartNum, OrdType, OrdPrefix, ContType, QtyInCont, QtyOfOrders, status);
                        loadGrid();
                    }
                }
                else
                {
                    string FlagNew = e.Cell.Row.Cells["NEW"].Text.ToString();
                    //   CheckBox chk=(CheckBox)(e.Cell.Row.Cells["NEW"]);
                    //  string status = e.Cell.Row.Cells["Tstatus"].Text.ToString();
                    string site_code = e.Cell.Row.Cells["TSiteCode"].Text.ToString();
                    string part_num = e.Cell.Row.Cells["TPartNum"].Text.ToString();
                    string displayPartNum = e.Cell.Row.Cells["TDisplayPN"].Text.ToString();
                    string OrdType = e.Cell.Row.Cells["TOrdType"].Text.ToString();
                    string OrdPrefix = e.Cell.Row.Cells["TOrdPrefix"].Text.ToString();
                    string ContType = e.Cell.Row.Cells["TContType"].Text.ToString();
                    int QtyInCont = Convert.ToInt32(e.Cell.Row.Cells["TQtyInCont"].Text.ToString());
                    int QtyOfOrders = Convert.ToInt32(e.Cell.Row.Cells["TQtyOfOrders"].Text.ToString());
                    string sqlInsert = " insert into [dbo].[KanbanMasterTemp]([TSiteCode],[TPartNum],[TDisplayPN],[TOrdType],[TOrdPrefix],[TContType],[TQtyInCont],[TQtyOfOrders],[TStatus],[TDelete]) values ('" + site_code + "','" + part_num + "','" + displayPartNum + "','" + OrdType + "','" + OrdPrefix + "','" + ContType + "','" + QtyInCont + "','" + QtyOfOrders + "','N','N')  Select @@identity  ";
                    DataTable dt = getdataSet(sqlInsert).Tables[0];
                    if (dt != null)
                        if (dt.Rows.Count > 0)
                            if (dt.Rows[0][0].ToString() != "0")
                                MessageBox.Show("Save Successfully");
                    loadGrid();
                }*/
        }

        public void DeleteRow(int id, string status)
        {
            //   string sqlUpdate = "exec sp_app_delete_kanbanMaster " + id.ToString() + ",'" + status + "'";
            //    getdataSet(sqlUpdate);

        }


        private void btnUpdate_Click(object sender, EventArgs e)
        {
            /*   string idString = "";
               foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.gridKanban.Rows)
               {
                   // MessageBox.Show(row.Cells[row.Cells.Count - 1].Text.ToString());
                   if (row.Cells["status"].Text.ToString() == "True")
                   {
                       idString = idString + "," + row.Cells["MasterID"].Value.ToString();
                   }
               }*/
            string sqlUpdate = " exec sp_app_updateKanbanProd 'KANBAN'";// + idString.ToString();
            DataTable dtSave = getdataSet(sqlUpdate).Tables[0];
            if (dtSave != null)
                if (dtSave.Rows.Count > 0)
                    if (dtSave.Rows[0][0].ToString() == "0")
                        MessageBox.Show("Production Kanban has been updated");
        }

        private void btnFilter_Click(object sender, EventArgs e)
        {
            loadGrid();
        }

        private void btnAddNew_Click(object sender, EventArgs e)
        {
            this.gridKanban.DisplayLayout.Bands[0].AddNew();
            this.gridKanban.Rows[gridKanban.Rows.Count - 1].Selected = true;
            this.gridKanban.Rows[gridKanban.Rows.Count - 1].Cells["TStatus"].Value = "N";
            this.gridKanban.Rows[gridKanban.Rows.Count - 1].Cells["TStatus"].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
            this.gridKanban.Rows[gridKanban.Rows.Count - 1].Cells["TDELETE"].Value = "False";
            this.gridKanban.Rows[gridKanban.Rows.Count - 1].Cells["TDELETE"].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
            this.gridKanban.Rows[gridKanban.Rows.Count - 1].Cells["NEW"].Value = "True";
        }

        private void button1_Click(object sender, EventArgs e)
        {

            String sql = " update KanbanMasterTemp set tstatus='N' ";
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.gridKanban.Rows)
            {
                int id = 0;
                if (row.Cells["MasterID"].Text.ToString().Trim() != "")
                    id = Convert.ToInt32(row.Cells["MasterID"].Text.ToString());
                string site_code = row.Cells["TSiteCode"].Text.ToString();
                string part_num = row.Cells["TPartNum"].Text.ToString();
                string displayPartNum = row.Cells["TDisplayPN"].Text.ToString();
                string OrdType = row.Cells["TOrdType"].Text.ToString();
                string OrdPrefix = row.Cells["TOrdPrefix"].Text.ToString();
                string ContType = row.Cells["TContType"].Text.ToString();
                string DeleteFlag = "N";
                if (row.Cells["TDelete"].Text.ToString().ToUpper() == "TRUE")
                    DeleteFlag = "Y";

                int QtyInCont = 0;
                int QtyOfOrders = 0;
                int MoqQtyOrders = 0;

                if (row.Cells["TQtyInCont"].Text.Trim().ToString() != "")
                    QtyInCont = Convert.ToInt32(row.Cells["TQtyInCont"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                if (row.Cells["TQtyOfOrders"].Text.Trim().ToString() != "")
                    QtyOfOrders = Convert.ToInt32(row.Cells["TQtyOfOrders"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                if (row.Cells["TMoqQty"].Text.Trim().ToString() != "")
                    MoqQtyOrders = Convert.ToInt32(row.Cells["TMoqQty"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());

                if (site_code != "")
                {
                    if (id == 0)
                    {
                        if (part_num != "")
                            sql = sql + "   insert into [dbo].[KanbanMasterTemp]([TSiteCode],[TPartNum],[TDisplayPN],[TOrdType],[TOrdPrefix],[TContType],[TQtyInCont],[TQtyOfOrders],[TStatus],[TDelete],TMoqQty) values ('" + site_code + "','" + part_num + "','" + displayPartNum + "','KANBAN','" + OrdPrefix + "','" + ContType + "','" + QtyInCont + "','" + QtyOfOrders + "','N','N'," + MoqQtyOrders + " )  ";
                    }
                    else
                    {
                        sql = sql + "  Update KanbanMasterTemp set TSiteCode='" + site_code + "',TPartNum='" + part_num + "',TDisplayPN='" + displayPartNum + "', TOrdType='" + OrdType + "', TOrdPrefix='" + OrdPrefix + "', TContType='" + ContType + "', TQtyInCont='" + QtyInCont + "', TQtyOfOrders='" + QtyOfOrders + "', TMoqQty='" + MoqQtyOrders + "', TDelete='" + DeleteFlag + "', TStatus='N'  where MasterID=" + id;
                    }
                }
            }
            if (sql != "")
            {

                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlTransaction tran = cn.BeginTransaction("Transaction");
                SqlCommand cmd = new SqlCommand(sql, cn);
                cmd.Transaction = tran;
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    tran.Commit();
                    MessageBox.Show("Updated successfully");
                }

                catch (Exception exp)
                {
                    string error = exp.Message.ToString();
                    tran.Rollback();
                    MessageBox.Show("Failed to Update. Please check the data");
                }
                cn.Close();
            }
            loadGrid();
        }

        private void gridKanban_AfterHeaderCheckStateChanged(object sender, Infragistics.Win.UltraWinGrid.AfterHeaderCheckStateChangedEventArgs e)
        {



        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            /*  DataTable dt = (this.gridKanban.DataSource as DataTable).Copy();
              //  RowsCollection rc ;
              foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in gridKanban.Rows)
              {
                  // MessageBox.Show(row.Cells[row.Cells.Count - 1].Text.ToString());
                  if (row.Cells[row.Cells.Count - 1].Text.ToString() == "True")
                  {
                      //  row.Delete();
                      DataRow[] drRep = dt.Select("MasterID=" + row.Cells["MasterID"].Value.ToString() + " and [part num]='" + row.Cells["part num"].Value.ToString() + "'");
                      if (drRep.Length > 0)
                          dt.Rows.Remove(drRep[0]);
                  }
              }

              gridKanban.DataSource = dt;  */
        }

        private void gridKanban_ClickCell(object sender, Infragistics.Win.UltraWinGrid.ClickCellEventArgs e)
        {
            if (e.Cell.Column.Index == 21)
            {
                if (e.Cell.Row.Cells["Masterid"].Value.ToString() != "")
                {
                    DeleteRowDataBase(Convert.ToInt32(e.Cell.Row.Cells["Masterid"].Value));
                }

                this.gridKanban.Rows[e.Cell.Row.Index].Delete();
            }
        }

        public void DeleteRowDataBase(int id)
        {
            string sqlDelete = "delete from kanbanmastertemp where masterid=" + id.ToString();
            getdataSet(sqlDelete);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            loadReplenishGrid();
        }

        private void btnUpdRep_Click(object sender, EventArgs e)
        {
            string sqlUpdate = " exec sp_app_updateKanbanProd 'REPLENISH' ";// + idString.ToString();
            DataTable dtSave = getdataSet(sqlUpdate).Tables[0];
            if (dtSave != null)
                if (dtSave.Rows.Count > 0)
                    if (dtSave.Rows[0][0].ToString() == "0")
                        MessageBox.Show("Production Replenish has been updated");

        }

        private void gridReplenish_ClickCell(object sender, Infragistics.Win.UltraWinGrid.ClickCellEventArgs e)
        {

            if (e.Cell.Column.Index == 20)
            {
                if (e.Cell.Row.Cells["TOrdType"].Value.ToString() != "KANBAN")
                {

                    if (e.Cell.Row.Cells["Masterid"].Value.ToString() != "")
                    {

                        DeleteRowDataBase(Convert.ToInt32(e.Cell.Row.Cells["Masterid"].Value));
                    }

                    this.gridReplenish.Rows[e.Cell.Row.Index].Delete();
                }
            }


        }

        private void btnAddRep_Click(object sender, EventArgs e)
        {

            this.gridReplenish.DisplayLayout.Bands[0].AddNew();
            this.gridReplenish.Rows[gridReplenish.Rows.Count - 1].Selected = true;
            this.gridReplenish.Rows[gridReplenish.Rows.Count - 1].Cells["TStatus"].Value = "N";
            this.gridReplenish.Rows[gridReplenish.Rows.Count - 1].Cells["TStatus"].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
            this.gridReplenish.Rows[gridReplenish.Rows.Count - 1].Cells["TDELETE"].Value = "False";
            this.gridReplenish.Rows[gridReplenish.Rows.Count - 1].Cells["TDELETE"].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
            this.gridReplenish.Rows[gridReplenish.Rows.Count - 1].Cells["NEW"].Value = "True";


        }

        private void btnSaveRep_Click(object sender, EventArgs e)
        {
            String sql = " update KanbanMasterTemp set tstatus='N' ";

            int id = 0;
            if (this.textRepID.Text.ToString().Trim() != "")
                id = Convert.ToInt32(textRepID.Text.ToString());
            string site_code = this.comboRepSite.Text.ToString();
            string part_num = this.textRepPartNum.Text.ToString();
            string displayPartNum = this.textRepDisplayPN.Text.ToString();
            string OrdType = ""; //= comboRepOrdType.Text.ToString();
            //   string OrdPrefix = row.Cells["TOrdPrefix"].Text.ToString();
            //    string ContType = row.Cells["TContType"].Text.ToString();
            string DeleteFlag = "N";
          /*  if (this.chkRepDelete.Text.ToString().ToUpper() == "TRUE")
                DeleteFlag = "Y";
            */

            int QtyInCont = 0;
            int QtyOfOrders = 0;
           
          
            QtyOfOrders = Convert.ToInt32(this.textRepQTYOrd.Value.ToString());
          
            if (id == 0)
            {
                if (part_num != "")
                    sql = sql + "   insert into [dbo].[KanbanMasterTemp]([TSiteCode],[TPartNum],[TDisplayPN],[TOrdType],[TQtyInCont],[TQtyOfOrders],[TStatus],[TDelete]) values ('" + site_code + "','" + part_num + "','" + displayPartNum + "','REPLENISH','" + QtyInCont + "','" + QtyOfOrders + "','N','N')  ";
            }
            else
            {
                sql = sql + "  Update KanbanMasterTemp set TSiteCode='" + site_code + "',TPartNum='" + part_num + "',TDisplayPN='" + displayPartNum + "', TOrdType='" + OrdType + "', TQtyInCont='" + QtyInCont + "', TQtyOfOrders='" + QtyOfOrders + "', TDelete='" + DeleteFlag + "', TStatus='N'  where MasterID=" + id;
            }

            if (sql != "")
            {

                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlTransaction tran = cn.BeginTransaction("Transaction");
                SqlCommand cmd = new SqlCommand(sql, cn);
                cmd.Transaction = tran;
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    tran.Commit();
                    MessageBox.Show("Updated successfully");

                }
                catch (Exception exp)
                {
                    string error = exp.Message.ToString();
                    tran.Rollback();
                    MessageBox.Show("Failed to Update. Please check the data");
                }
                cn.Close();
            }
            loadReplenishGrid();
        }

        private void gridReplenish_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
        {
            if ((e.Cell.Column.Index == 6) || (e.Cell.Column.Index == 5))
            {
                int QtyInCont = 0;
                int QtyOfOrders = 0;
                if (e.Cell.Row.Cells["TQtyInCont"].Text.ToString() != "")
                {
                    QtyInCont = Convert.ToInt32(e.Cell.Row.Cells["TQtyInCont"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                }
                if (e.Cell.Row.Cells["TQtyOfOrders"].Text.ToString() != "")
                {
                    QtyOfOrders = Convert.ToInt32(e.Cell.Row.Cells["TQtyOfOrders"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                }
                int tot_qty = QtyInCont * QtyOfOrders;

                e.Cell.Row.Cells["Tot_Qty"].Value = tot_qty;
            }
        }

        private void gridKanban_KeyDown(object sender, KeyEventArgs e)
        {

            Infragistics.Win.UltraWinGrid.UltraGridCell oldIndex = null;

            switch (e.KeyCode)
            {
                case Keys.Enter:
                    oldIndex = gridKanban.ActiveCell;
                    gridKanban.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, false, false);
                    gridKanban.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.BelowRow, false, false);
                    gridKanban.ActiveCell = gridKanban.ActiveRow.Cells[oldIndex.Column];
                    e.Handled = true;
                    gridKanban.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, false, false);
                    break;

                case Keys.Down:
                    oldIndex = gridKanban.ActiveCell;
                    gridKanban.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, false, false);
                    gridKanban.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.BelowRow, false, false);
                    gridKanban.ActiveCell = gridKanban.ActiveRow.Cells[oldIndex.Column];
                    e.Handled = true;
                    gridKanban.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, false, false);
                    break;

                case Keys.Up:
                    oldIndex = gridKanban.ActiveCell;
                    gridKanban.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, false, false);
                    gridKanban.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.AboveRow, false, false);
                    gridKanban.ActiveCell = gridKanban.ActiveRow.Cells[oldIndex.Column];
                    e.Handled = true;
                    gridKanban.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, false, false);
                    break;
            }

        }

        private void gridReplenish_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
        {

            if ((e.Cell.Column.Index == 6) || (e.Cell.Column.Index == 7))
            {
                int QtyInCont = 0;
                int QtyOfOrders = 0;
                if (e.Cell.Row.Cells["TQtyInCont"].Text.ToString() != "")
                {
                    QtyInCont = Convert.ToInt32(e.Cell.Row.Cells["TQtyInCont"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                }
                if (e.Cell.Row.Cells["TQtyOfOrders"].Text.ToString() != "")
                {
                    QtyOfOrders = Convert.ToInt32(e.Cell.Row.Cells["TQtyOfOrders"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                }
                int tot_qty = QtyInCont * QtyOfOrders;
                e.Cell.Row.Cells["Tot_Qty"].Value = tot_qty;
            }

        }
        string file = "";
        string foldername = "\\\\mvfile\\MV Reports\\MAster Data & forms\\PlanningExecutionForms\\Templates\\";
        public void create_file_OLEDB()
        {
            Microsoft.Office.Interop.Excel.Workbook wk;
            string strcon = @"Provider=SQLOLEDB;Data Source=MVerp; Trusted_Connection=yes;";

            System.Data.OleDb.OleDbConnection cn = new System.Data.OleDb.OleDbConnection(strcon);
            //   cn.Open();
            Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application();


            string file_path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\";
            string file_name = "Kanban Master_" + System.DateTime.Now.Month + System.DateTime.Now.Day + System.DateTime.Now.Year + System.DateTime.Now.Hour + System.DateTime.Now.Minute + "-" + System.DateTime.Now.Second.ToString();
            //  create_file();
            file = file_path + file_name + ".xlsx";
            //   if (System.IO.File.Exists(file))
            //  System.IO.File.Delete(file);
            file = file.Replace("file:\\", "");

            String workbookPath = foldername +"Kanban Update for Shared Doc.xlsx";
            wk = excelapp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            wk.SaveAs(file, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wk.Sheets[1];

          //  String stsql1 = "  SELECT     * FROM   mimdist.dbo.ExcelKanbanCompare AS ExcelKanbanCompare ORDER BY SiteCode, PartNum";
            String stsql1 = "  SELECT   *   FROM   mimdist.dbo.[GDOC_KanbanCompare] AS ExcelKanbanCompare ORDER BY SiteCode, PartNum";
            //  Object txt = new Object();

            try
            {
                Microsoft.Office.Interop.Excel.QueryTable oQryTable;//=new Microsoft.Office.Interop.Excel.QueryTable();
                oQryTable = ws.QueryTables.Add("OLEDB;" + strcon, ws.Range["A1"], stsql1);
                oQryTable.RefreshStyle = Microsoft.Office.Interop.Excel.XlCellInsertionMode.xlOverwriteCells;
                oQryTable.Refresh(false);
                wk.Save();
                //   TimeSpan interval = new TimeSpan(0,1, 2);
                //  System.Threading.Thread.Sleep(interval);
                Microsoft.Office.Interop.Excel.Worksheet ws1 = (Microsoft.Office.Interop.Excel.Worksheet)wk.Sheets[2];

                String stsql2 = " SELECT    *  FROM   mimdist.dbo.GDOC_ReplenishmentMaster AS ExcelReplenishMaster  ORDER BY SiteCode,orderType, PartNum ";          //  Object txt = new Object();


                Microsoft.Office.Interop.Excel.QueryTable oQryTable1;//=new Microsoft.Office.Interop.Excel.QueryTable();
                oQryTable1 = ws1.QueryTables.Add("OLEDB;" + strcon, ws1.Range["A1"], stsql2);
                oQryTable1.RefreshStyle = Microsoft.Office.Interop.Excel.XlCellInsertionMode.xlOverwriteCells;
                oQryTable1.Refresh(false);
                wk.Save();


                Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)wk.Sheets[3];

                String stsql3 = " SELECT    *  FROM         mimdist.dbo.GDOC_ReplenDynamic AS ExcelReplenishMaster  ORDER BY SiteCode,orderType, PartNum ";          //  Object txt = new Object();


                Microsoft.Office.Interop.Excel.QueryTable oQryTable2;//=new Microsoft.Office.Interop.Excel.QueryTable();
                oQryTable2 = ws2.QueryTables.Add("OLEDB;" + strcon, ws2.Range["A1"], stsql3);
                oQryTable2.RefreshStyle = Microsoft.Office.Interop.Excel.XlCellInsertionMode.xlOverwriteCells;
                oQryTable2.Refresh(false);
                wk.Save();
            }
            catch (Exception ex)
            {
                //  MessageBox.Show(ex.Message.ToString())
            }
            finally
            {
                wk.Close(true, workbookPath, Type.Missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wk);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp);
            }
            // System.Diagnostics.Process p = new System.Diagnostics.Process();
            // p.StartInfo.FileName = file_name;
            // p.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized;
            // p.Start();
        }

        private void ultraButton1_Click(object sender, EventArgs e)
        {
            
            create_file_OLEDB();
            uploadFile();
        }

     

        string collectionName = "";
        public void uploadFile()
        {

                   
     
                string fileid = "";



                UserCredential credential;

                using (var stream =
                    new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = System.Environment.GetFolderPath(
                        System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets, new[] { DriveService.Scope.Drive }
                       ,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Drive API service.
                service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "MIM Project",
                });

        fileid = getGIDByName("Vanilla KanBan Parts List update ");
        Google.Apis.Drive.v3.Data.File body = service.Files.Get(fileid).Execute();
         
         

                try
                {
                    Google.Apis.Drive.v3.Data.File body1 = new Google.Apis.Drive.v3.Data.File();
                    body.Name = "Vanilla KanBan Parts List update ";
                     body.Description = "Vanilla KanBan Parts List update ";
                       body.MimeType = "application/xlsx";                    
                    string filename = file;
                    System.IO.Stream fileStream = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                    byte[] byteArray = new byte[fileStream.Length];
                    fileStream.Read(byteArray, 0, (int)fileStream.Length);
                    System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
                    // FilesResource.InsertMediaUpload request = dr.Files.Insert(body1,stream, "application/x-vnd.oasis.opendocument.spreadsheet");    
                    FilesResource.UpdateMediaUpload request = service.Files.Update(body, fileid, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
                    //request.Convert = true;
                   request.Upload();
                    MessageBox.Show("File has been uploaded in shared doc");
                    fileStream.Close();
                    fileStream.Dispose();
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Exception to upload file");
                }
                finally
                {

                }
            }
       
        private void ultraButton1_Click_1(object sender, EventArgs e)
        {
            //connectSharedDoc();
            //create_file_OLEDB();
            //uploadFile();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void ultraTextEditor1_Leave(object sender, EventArgs e)
        {
            if (this.textPartNum.Text.Trim() != "")
            {
                if (!validatePart(this.textPartNum.Text.Trim()))
                {
                    this.textPartNum.Text = "";
                    this.textPartNum.Focus();
                }
            }
            else
            {
                MessageBox.Show("Please enter part num");
                this.textPartNum.Focus();
            }
        }

        public Boolean validatePart(string part_num)
        {
            string sql = "  Select count(*) from inv_master where part_no='" + part_num + "'";
            DataSet ds = getdataSet(sql);
            if (ds.Tables[0].Rows[0][0].ToString() == "0")
            {
                MessageBox.Show("Part doesn't exists! Please check");
                return false;
            }
            return true;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            clearKanban();
        }

        public void clearKanban()
        {
            this.textKanID.Text = "";
            this.textMoqQty.Value = 0;
            this.textDisplayPN.Text = "";
            this.textPartNum.Text = "";
           // this.comboPrefix.Text = "";
          //  this.comboContType.Text = "";
            this.comboSite.Text = "";
            this.textQtyCont.Value = 0;
            this.textQtyOrd.Value = 0;
         //   this.chkMarkAsdelete.Checked = false;
            this.chkOnhold.Checked = false;
            this.chkTLAFlag.Checked = false;


            foreach (Control ctrl in this.pnlContiner.Controls)
            {
                if (ctrl is RadioButton)
                {
                    RadioButton rad = ctrl as RadioButton;
                         rad.Checked = false;
                       
                }
            }

            foreach (Control ctrl in this.pnlOrderPrefix.Controls)
            {
                if (ctrl is RadioButton)
                {
                    RadioButton rad = ctrl as RadioButton;
                    rad.Checked = false;
                }
            }
         

          

            
        }


        private void gridKanban_AfterSelectChange(object sender, Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs e)
        {
            if (this.gridKanban.Selected.Rows.Count > 0)
            {
                this.textKanID.Text = gridKanban.Selected.Rows[0].Cells["MasterID"].Value.ToString();
                this.textMoqQty.Value = gridKanban.Selected.Rows[0].Cells["TMoqQty"].Value.ToString();
                this.textDisplayPN.Text = gridKanban.Selected.Rows[0].Cells["TDisplayPN"].Value.ToString();
                this.textPartNum.Text = gridKanban.Selected.Rows[0].Cells["TPartNum"].Value.ToString();
              //  this.comboPrefix.Text = gridKanban.Selected.Rows[0].Cells["TOrdPrefix"].Value.ToString();
            //    this.comboContType.Text = gridKanban.Selected.Rows[0].Cells["TContType"].Value.ToString();
                this.comboSite.Text = gridKanban.Selected.Rows[0].Cells["TSiteCode"].Value.ToString();
                this.textQtyCont.Value = gridKanban.Selected.Rows[0].Cells["TQtyInCont"].Value.ToString();
                this.textQtyOrd.Value = gridKanban.Selected.Rows[0].Cells["TQtyOfOrders"].Value.ToString();


                foreach (Control ctrl in this.pnlOrderPrefix.Controls)
                {
                    if (ctrl is RadioButton)
                    {
                        RadioButton rad = ctrl as RadioButton;
                        if (rad.Name == gridKanban.Selected.Rows[0].Cells["TOrdPrefix"].Value.ToString())
                        {

                            rad.Checked = true;
                            break;
                        }
                    }
                }
                string ContType = "";

                foreach (Control ctrl in this.pnlContiner.Controls)
                {
                    if (ctrl is RadioButton)
                    {
                        RadioButton rad = ctrl as RadioButton;
                        if (rad.Name == gridKanban.Selected.Rows[0].Cells["TContType"].Value.ToString())
                        {

                            rad.Checked = true;
                            break;
                        }
                    }
                }


                if (gridKanban.Selected.Rows[0].Cells["TOnHold"].Value.ToString() == "True")
                    this.chkOnhold.Checked = true;
                else
                    this.chkOnhold.Checked = false;

                if (gridKanban.Selected.Rows[0].Cells["TTLAFlag"].Value.ToString() == "True")
                    this.chkTLAFlag.Checked = true;
                else
                    this.chkTLAFlag.Checked = false;

             //   if (gridKanban.Selected.Rows[0].Cells["TDelete"].Value.ToString() == "True")
             //       this.chkMarkAsdelete.Checked = true;
             //   else
                //    this.chkMarkAsdelete.Checked = false;
               
               
            }
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            String sql = " update KanbanMasterTemp set tstatus='N' ";

            int id = 0;
            if (this.textKanID.Text != "")
                id = Convert.ToInt32(this.textKanID.Text);
            string site_code = comboSite.Text.ToString();
            string part_num = this.textPartNum.Text;
            string displayPartNum = this.textDisplayPN.Text;
            string OrdType = "KANBAN";

            string OrdPrefix = "";

            foreach (Control ctrl  in this.pnlOrderPrefix.Controls)
            {
                if (ctrl is RadioButton )
                    {
                        RadioButton rad = ctrl as RadioButton;
                        if ( rad.Checked) 
                        {
                            OrdPrefix=  rad.Text;
                            break;
                        }
                    }
            }
            string ContType = "";

            foreach (Control ctrl in this.pnlContiner.Controls)
            {
                if (ctrl is RadioButton)
                {
                    RadioButton rad = ctrl as RadioButton;
                    if (rad.Checked)
                    {
                        ContType = rad.Text;
                        break;
                    }
                }
            }
          

       
         // ******// if (this.comboContType.Text!="SELECT")

            // ******//     ContType= this.comboContType.Text;
         //   string DeleteFlag = "N";
            string Onhold = "N";
            string tlaFlag = "N";
        //    if (chkMarkAsdelete.Checked.ToString().ToUpper() == "TRUE")
          //      DeleteFlag = "Y";

            if (chkOnhold.Checked.ToString().ToUpper() == "TRUE")
                Onhold = "Y";

            if (this.chkTLAFlag.Checked.ToString().ToUpper() == "TRUE")
                tlaFlag = "Y";

            int QtyInCont = 0;
            int QtyOfOrders = 0;
            int MoqQtyOrders = 0;

            if (this.textQtyCont.Value.ToString() != "")
                QtyInCont = Convert.ToInt32(this.textQtyCont.Value);


            if (this.textQtyOrd.Value.ToString() != "")
                QtyOfOrders = Convert.ToInt32(this.textQtyOrd.Value);

            if (this.textMoqQty.Value.ToString() != "")
                MoqQtyOrders = Convert.ToInt32(this.textMoqQty.Value);

            if (ContType == "")
                MessageBox.Show("Select Container type");
            else
            {
            if (site_code != "")
            {
                if (id == 0)
                {
                    if (part_num != "")
                        sql = sql + "   insert into [dbo].[KanbanMasterTemp]([TSiteCode],[TPartNum],[TDisplayPN],[TOrdType],[TOrdPrefix],[TContType],[TQtyInCont],[TQtyOfOrders],[TStatus]"/*,[TDelete]*/+ ",TMoqQty,TonHold) values ('" + site_code + "','" + part_num + "','" + displayPartNum + "','KANBAN','" + OrdPrefix + "','" + ContType + "','" + QtyInCont + "','" + QtyOfOrders + "','N'," + MoqQtyOrders + ",'N' )  ";//'N',
                }
                else
                {
                    sql = sql + "  Update KanbanMasterTemp set TSiteCode='" + site_code + "',TPartNum='" + part_num + "',TDisplayPN='" + displayPartNum + "', TOrdType='" + OrdType + "', TOrdPrefix='" + OrdPrefix + "', TContType='" + ContType + "', TQtyInCont='" + QtyInCont + "', TQtyOfOrders='" + QtyOfOrders + /*"', TDelete='" + DeleteFlag +*/ "', TTLAFlag='" + tlaFlag + "', TOnHold='" + Onhold + "', TStatus='N'  where MasterID=" + id;
                }
            }

        }

            if (sql != " update KanbanMasterTemp set tstatus='N' ")
            {

                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlCommand cmd = new SqlCommand(sql, cn);
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    //  tran.Commit();
                    MessageBox.Show("Updated successfully");
                }

                catch (Exception exp)
                {
                    string error = exp.Message.ToString();
                    // tran.Rollback();
                    MessageBox.Show("Failed to Update. Please check the data");
                }
                cn.Close();
            }
            loadGrid();
            clearKanban();
        }

        private void gridReplenish_AfterSelectChange(object sender, Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs e)
        {
            if (this.gridReplenish.Selected.Rows.Count > 0)
            {
                txtOrderType.Text = gridReplenish.Selected.Rows[0].Cells["TOrdType"].Value.ToString();
                this.txtOrderType.Visible = true;
                label20.Visible = true;
                if (gridReplenish.Selected.Rows[0].Cells["TOrdType"].Value.ToString() == "REPLENISH")
                {
                    
                    this.textRepQTYMoq.Enabled=true;
                    this.textRepDisplayPN.Enabled = true;
                    this.textRepPartNum.Enabled = true;
                    this.comboRepSite.Enabled = true;
                    this.textRepQTYOrd.Enabled = true;
                    textRepQTYOrd.Enabled = true;
                }
                else
                {
                    textRepQTYOrd.Enabled = false;
                    this.textRepQTYMoq.Enabled = true;
                    this.textRepDisplayPN.Enabled = false;
                    this.textRepPartNum.Enabled = false;
                    this.comboRepSite.Enabled = false;
                  
                }

                    this.textRepID.Text = gridReplenish.Selected.Rows[0].Cells["MasterID"].Value.ToString();
                    this.textRepQTYMoq.Value = gridReplenish.Selected.Rows[0].Cells["TMoqQty"].Value.ToString();
                    this.textRepDisplayPN.Text = gridReplenish.Selected.Rows[0].Cells["TDisplayPN"].Value.ToString();
                    this.textRepPartNum.Text = gridReplenish.Selected.Rows[0].Cells["TPartNum"].Value.ToString();                  
                    this.comboRepSite.Text = gridReplenish.Selected.Rows[0].Cells["TSiteCode"].Value.ToString();                 
                    this.textRepQTYOrd.Value = gridReplenish.Selected.Rows[0].Cells["TSetQTY"].Value.ToString();

                    if (gridReplenish.Selected.Rows[0].Cells["TOnHold"].Value.ToString().ToUpper() == "TRUE")
                        this.chkRepOnhold.Checked = true;
                    else
                        this.chkRepOnhold.Checked = false;

                    if (gridReplenish.Selected.Rows[0].Cells["TTLAFlag"].Value.ToString().ToUpper() == "TRUE")
                        this.chkRepTLAFlag.Checked = true;
                    else
                        this.chkRepTLAFlag.Checked = false;

                 //   if (gridReplenish.Selected.Rows[0].Cells["TDelete"].Value.ToString().ToUpper() == "TRUE")
                  //      this.chkRepDelete.Checked = true;
                //    else
                //        this.chkRepDelete.Checked = false;
               


               // }
               // else
               // {
                //    clearReplishment();

                //}
            }
        }

        private void btnRepSave_Click(object sender, EventArgs e)
        {
            String sql = " update KanbanMasterTemp set tstatus='N' ";


            int id = 0;
            if (this.textRepID.Text.ToString().Trim() != "")
                id = Convert.ToInt32(this.textRepID.Text.ToString());
            string site_code = this.comboRepSite.Text.ToString();
            string part_num = this.textRepPartNum.Text.ToString();
            string displayPartNum = this.textRepDisplayPN.Text.ToString();


         //   string DeleteFlag = "N";
            string tlaFlag = "N";
         //   if (this.chkRepDelete.Checked.ToString().ToUpper() == "TRUE")
          //      DeleteFlag = "Y";
            string OnHold = "N";

            if (this.chkRepOnhold.Checked.ToString().ToUpper() == "TRUE")
                OnHold = "Y";

            if (this.chkRepTLAFlag.Checked.ToString().ToUpper() == "TRUE")
                tlaFlag = "Y";

            int MoqQTY = 0;
            int QtyOfOrders = 0;
            // if (row.Cells["TQtyInCont"].Text.ToString() != "")
            // {
            MoqQTY = Convert.ToInt32(this.textRepQTYMoq.Value.ToString());
            // }
            //  if (row.Cells["TQtyOfOrders"].Text.ToString() != "")
            //  {
            QtyOfOrders = Convert.ToInt32(this.textRepQTYOrd.Value.ToString());
            //  }
            if (id == 0)
            {
                if (part_num != "")
                    sql = sql + "   insert into [dbo].[KanbanMasterTemp]([TSiteCode],[TPartNum],[TDisplayPN],[TOrdType],TMoqQty,[TQtyOfOrders],[TStatus]"+/*,[TDelete]*/",[TOnHold]) values ('" + site_code + "','" + part_num + "','" + displayPartNum + "','REPLENISH','" + MoqQTY + "','" + QtyOfOrders + "','N','N')  ";//,'N'
            }
            else
            {
                if (txtOrderType.Text=="REPLENISH")
                    sql = sql + "  Update KanbanMasterTemp set TSiteCode='" + site_code + "',TPartNum='" + part_num + "',TDisplayPN='" + displayPartNum + "', TMoqQty='" + MoqQTY + "', TQtyOfOrders='" + QtyOfOrders +/* "', TDelete='" + DeleteFlag +*/ "', TTLAFlag='" + tlaFlag + "', TOnHold='" + OnHold + "', TStatus='N'  where MasterID=" + id;
                else
                    sql = sql + "  Update KanbanMasterTemp set  TMoqQty='" + MoqQTY + "'  where MasterID=" + id;
            }

            if (sql != "")
            {
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlTransaction tran = cn.BeginTransaction("Transaction");
                SqlCommand cmd = new SqlCommand(sql, cn);
                cmd.Transaction = tran;
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    tran.Commit();
                    MessageBox.Show("Updated successfully");
                }
                catch (Exception exp)
                {
                    string error = exp.Message.ToString();
                    tran.Rollback();
                    MessageBox.Show("Failed to Update. Please check the data");
                }
                cn.Close();
            }

            loadReplenishGrid();
        }

        public void clearReplishment()
        {
            this.textRepID.Text = "";
            this.textRepQTYMoq.Value = 0;
            this.textRepDisplayPN.Text = "";
            this.textRepPartNum.Text = "";
                 
            this.comboRepSite.Text = "";
            this.textRepQTYOrd.Value = 0;


            this.textRepQTYMoq.Enabled = true;
            this.textRepDisplayPN.Enabled = true;
            this.textRepPartNum.Enabled = true;

            this.comboRepSite.Enabled = true;
            this.textRepQTYOrd.Enabled = true;
            

            label20.Visible = false;
          //  this.chkRepDelete.Checked = false;
            this.chkRepOnhold.Checked = false;
            this.chkRepTLAFlag.Checked = false;
            textRepQTYOrd.Enabled = true;
            txtOrderType.Visible = false;

        }


        private void btnRepAddNew_Click(object sender, EventArgs e)
        {
            clearReplishment();
        }

        private void btnRepCancel_Click(object sender, EventArgs e)
        {

        }

        private void comboRepOrdType_ValueChanged(object sender, EventArgs e)
        {
            //if (this.comboRepOrdType.Text == "KANBAN")
            //{
            //    this.comboRepOrdPrefix.Visible = true;
            //    this.comboRepContType.Visible = true;
            //    this.lblContType.Visible = true;
            //    this.lblOrdPrefix.Visible = true;
            //}
            //else
            //{
            //    this.comboRepOrdPrefix.Visible = false;
            //    this.comboRepContType.Visible = false;
            //    this.lblContType.Visible = false;
            //    this.lblOrdPrefix.Visible = false;
            //}
        }

        private void textRepPartNum_Leave(object sender, EventArgs e)
        {
            if (this.textRepPartNum.Text.Trim() != "")
            {
                if (!validatePart(this.textRepPartNum.Text.Trim()))
                {
                    this.textRepPartNum.Text = "";
                    this.textRepPartNum.Focus();
                }
            }
            else
            {
                MessageBox.Show("Please enter part num");
                this.textRepPartNum.Focus();
            }
        }

        private void textQtyOrd_Leave(object sender, EventArgs e)
        {
            getTotalQty();
        }

        public void getTotalQty()
        {
            this.textMoqQty.Value = Convert.ToInt32(textQtyOrd.Value) * Convert.ToInt32(textQtyCont.Value);

        }

        private void textQtyCont_Leave(object sender, EventArgs e)
        {
            getTotalQty();
        }

        private void textQtyOrd_Layout(object sender, LayoutEventArgs e)
        {

        }
        DriveService service;
        string googleshareDoc;
        public void ImportsharedocToExcel()
        {

            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets, new[] { DriveService.Scope.Drive }
                   ,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Drive API service.
            service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "MIM Project",
            });

            string pageToken = null;
            string fileId = "";
            string filename = "Vanilla KanBan Parts List update ";

            fileId = getGIDByName(filename);
            ListViewItem item = new ListViewItem(new string[2] { filename, fileId });
            googleshareDoc = filename;

            saveTo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\VanillaKanBanParts.xlsx";
            saveTo = saveTo.Replace("file:\\", "");
            if (downloadfile(service, fileId, saveTo))
            {
                showInGrid("Replenishment List ");
                showInGrid("Dynamic List");
            }
            MessageBox.Show("Updated successfully");




        /*    string fileName = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\currentRackDeliveryPlan.xlsx";
            fileName = fileName.Replace("file:\\", "");

            Google.Apis.Drive.v3.Data.File file1 = service.Files.Get(fileId).Execute();
            if (this.downloadfile(service, fileId, fileName))
            {

                getallSheet(fileName);
                //  MessageBox.Show("File has been downloaded");

            }

            FilesResource.ListRequest lstreq = service.Files.List();
            do
            {
                try
                {
                    FileList fileList = lstreq.Execute();
                    foreach (Google.Apis.Drive.v2.Data.File fileitem in fileList.Items)
                    {
                        if (fileitem.Title=="Vanilla KanBan Parts List update ")
                        {
                            ListViewItem item = new ListViewItem(new string[2] { fileitem.Title, fileitem.Id });
                            googleshareDoc = fileitem.Title;

                            saveTo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\VanillaKanBanParts.xlsx";
                            saveTo = saveTo.Replace("file:\\", "");
                            if (downloadFile(service, fileitem, saveTo))
                            {
                                showInGrid("Replenishment List ");
                                showInGrid("Dynamic List");
                               }
                            MessageBox.Show("Updated successfully");
                            break;
                        }

                    }
                    lstreq.PageToken = fileList.NextPageToken;
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                    lstreq.PageToken = null;
                }
            } while (!String.IsNullOrEmpty(lstreq.PageToken));

            */

        }

     

        public string getGIDByName(string strName)
        {
            string parent_id = "";
            string pageToken = null;
            do
            {
                var request = service.Files.List();

                request.Q = "name = '" + strName + "'";//" + 

                //  string parent_id = spreadsheetListView.SelectedItems[0].SubItems[1].Text;
                // request.Q = "'" + parent_id + "' in parents";
                request.Spaces = "Drive";
                request.Fields = "nextPageToken, files(id, name)";
                request.PageToken = pageToken;

                var result = request.Execute();
                foreach (var file in result.Files)
                {
                    parent_id = file.Id;
                    break;
                }
                pageToken = result.NextPageToken;
            } while (pageToken != null);


            return parent_id;

        }

        string saveTo = "";


        public bool downloadfile(DriveService service, string fileId, string fileName)
        {

            var request = service.Files.Export(fileId, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            //   var fileId1 = "0BwwA4oUTeiV1UVNwOHItT0xfa2M";
            //  var request = service.Files.Get(fileId1);



            Stream streamvar = request.ExecuteAsStream();


            using (FileStream fs = System.IO.File.Create(fileName))
            {

                byte[] buffer = new byte[8 * 1024];
                int len;
                while ((len = streamvar.Read(buffer, 0, buffer.Length)) > 0)
                {
                    fs.Write(buffer, 0, len);
                }
                return true;
            }


            return false;
        }


   /*     public Boolean downloadFile(DriveService _service, File _fileResource, string _saveTo)
        {
            if (!String.IsNullOrEmpty(_fileResource.ExportLinks["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]))
            {
                try
                {
                    var x = _service.HttpClient.GetByteArrayAsync(_fileResource.ExportLinks["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]);
                    byte[] arrBytes = x.Result;

                    System.IO.File.WriteAllBytes(_saveTo, arrBytes);

                    return true;
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                    return false;
                }
            }
            else
            {
                // The file doesn't have any content stored on Drive.
                return false;
            }
        }*/

        string table_name = "";
        string ALLColumns = "";
        public void showInGrid(string table_name)
        {
            System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + saveTo + "';Extended Properties=Excel 12.0;");
            System.Data.OleDb.OleDbDataAdapter MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [" + table_name + "$]", MyConnection);
            MyCommand.TableMappings.Add("Table", "TestTable");
            System.Data.DataSet DtSet = new System.Data.DataSet();
            MyCommand.Fill(DtSet);
            DataTable dtShareColl = DtSet.Tables[0];
            MyConnection.Close();

            SetListListView(dtShareColl);

        }
        //public void connectSharedDoc()
        //{
        //    service = new SpreadsheetsService("service");
        //    service.setUserCredentials("admin@materialinmotion.com", "mimi1000");   //tedatmim@gmail.com ,"mimi100"     
        //    editUriTable = new System.Collections.Hashtable();
        //    SpreadsheetQuery qrySpreadsheet = new SpreadsheetQuery();
        //    SpreadsheetFeed sFeed = service.Query(qrySpreadsheet);

        //    bool bFound = false;

        //    AtomEntryCollection entries = sFeed.Entries;
        //    DataTable dtShareColl;
        //    for (int i = 0; i < entries.Count; i++)
        //    {
        //        if (entries[i].Title.Text == "Vanilla KanBan Parts List update")
        //        {
        //            // Get the worksheets feed URI
        //            AtomLink worksheetsLink = entries[i].Links.FindService(GDataSpreadsheetsNameTable.WorksheetRel, AtomLink.ATOM_TYPE);
        //            WorksheetQuery query = new WorksheetQuery(worksheetsLink.HRef.Content);
        //            WorksheetFeed wFeed = service.Query(query);
        //            SetWorksheetListView(wFeed);
        //        }
        //    }
        //}
        //string WorkSheet = "";
        //void SetWorksheetListView(WorksheetFeed feed)
        //{           
        //    AtomEntryCollection entries = feed.Entries;
        //    for (int i = 0; i < entries.Count; i++)
        //    {                
        //        // Get the list feed URI
        //        if (entries[i].Title.Text.ToString().Trim().ToUpper() == "REPLENISHMENT LIST")
        //        {
        //            AtomLink listLink = entries[i].Links.FindService(GDataSpreadsheetsNameTable.ListRel, AtomLink.ATOM_TYPE);
        //            ListItem1 lst = new ListItem1(listLink.HRef.ToString(), entries[i].Title.Text); // + " (Rows:" + ((WorksheetEntry)entries[i]).RowCount.Count + ", Cols:" + ((WorksheetEntry)entries[i]).ColCount.Count + ")"
        //            ListQuery query = new ListQuery(listLink.HRef.ToString());
        //            WorkSheet = lst.Name.ToString();
        //            SetListListView(query);
        //        }

        //        if (entries[i].Title.Text.ToString().Trim().ToUpper() == "DYNAMIC LIST")
        //        {
        //            AtomLink listLink = entries[i].Links.FindService(GDataSpreadsheetsNameTable.ListRel, AtomLink.ATOM_TYPE);
        //            ListItem1 lst = new ListItem1(listLink.HRef.ToString(), entries[i].Title.Text); // + " (Rows:" + ((WorksheetEntry)entries[i]).RowCount.Count + ", Cols:" + ((WorksheetEntry)entries[i]).ColCount.Count + ")"
        //            ListQuery query = new ListQuery(listLink.HRef.ToString());
        //            WorkSheet = lst.Name.ToString();
        //            SetListListView(query);
        //        }
        //    }
        //    MessageBox.Show("Updated successfully");
        //}

        private void button1_Click_3(object sender, EventArgs e)
        {
            //connectSharedDoc();

        }

        //public DataTable DownloadFileToDataGrid(ListQuery query, string plan_type)
        //{
        //    ListFeed lstFeed = this.service.Query(query);
        //    AtomEntryCollection entries = lstFeed.Entries;

        //    string sharedoccolumns="";

        //    DataTable    table = new DataTable();
        //    //  MessageBox.Show(entries.Count.ToString());
        //    for (int i = 0; i < entries.Count; i++)
        //    {
        //        ListEntry entry = entries[i] as ListEntry;
        //        ListViewItem item = new ListViewItem();
        //        if (entry != null)
        //        {
        //            ListEntry.CustomElementCollection elements = entry.Elements;
        //            // MessageBox.Show(elements.Count.ToString());
        //            if (table.Rows.Count < 1)
        //            {
        //                for (int j = 0; j < elements.Count; j++)
        //                {
        //                    // DataGridViewColumnHeaderCell cell = new DataGridViewColumnHeaderCell();
        //                    //MessageBox.Show(elements[j].LocalName);
        //                    table.Columns.Add(elements[j].LocalName.ToUpper(), typeof(string));
        //                    sharedoccolumns = sharedoccolumns + elements[j].LocalName.ToUpper() + ",";
        //                }
        //            }
        //            DataRow myDataRow = table.NewRow();
        //            for (int j = 0; j < elements.Count; j++)
        //            {
        //                if (j > table.Columns.Count - 1)
        //                {
        //                    table.Columns.Add("col" + j);
        //                }

        //                if (elements[j].Value != null)
        //                {
        //                    myDataRow[elements[j].LocalName] = elements[j].Value;
        //                }
        //                else
        //                {
        //                    myDataRow[j] = "";
        //                }
        //            }
        //            table.Rows.Add(myDataRow);
        //        }
        //    }
                           
        //    return table;
            
        //}

        static string ExtractNumbers(string expr)
        {
            return string.Join(null, System.Text.RegularExpressions.Regex.Split(expr, "[^\\d]"));
        }

        void SetListListView(DataTable dtShareColl)
        {
            
           // DataTable dtShareColl = DownloadFileToDataGrid(lstquery, WorkSheet);
            string updateQRY="";
            int MoqQTY = 0;
             int  setQty=0;
             for (int i = 0; i < dtShareColl.Rows.Count; i++)
             {
                 if (dtShareColl.Rows[i]["MasterID"].ToString() != "")
                 {
                     int MAsterID = Convert.ToInt32(dtShareColl.Rows[i]["MasterID"].ToString());
                     if (dtShareColl.Rows[i]["Upload New MOQ"].ToString() != "")
                     {
                         MoqQTY = Convert.ToInt32(ExtractNumbers(dtShareColl.Rows[i]["Upload New MOQ"].ToString()));
                     }

                     string OrderType = dtShareColl.Rows[i]["OrderType"].ToString();
                     if (dtShareColl.Rows[i]["SetQty"].ToString() != "")
                         setQty = Convert.ToInt32(ExtractNumbers(dtShareColl.Rows[i]["SetQty"].ToString()));
                     if (OrderType == "REPLENISH")
                         updateQRY = updateQRY + "    update dbo.KanbanMasterTemp set [TQtyOfOrders]=" + setQty + "  where [MasterID]=" + MAsterID + "  and TOrdType='REPLENISH'       update dbo.KanbanMaster set [QtyOfOrders]=" + setQty + "  where [MasterID]=" + MAsterID + "  and OrdType='REPLENISH' ";

                     updateQRY = updateQRY + "    update dbo.KanbanMasterTemp set [TMoqQty]=" + MoqQTY + "  where [MasterID]=" + MAsterID + "    update dbo.KanbanMaster set [MoqQty]=" + MoqQTY + "  where [MasterID]=" + MAsterID;
                 }
             }
            updateQRY = updateQRY + " select @@error  ";
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlTransaction tran = cn.BeginTransaction("Transaction");
            SqlCommand cmd = new SqlCommand(updateQRY, cn);        
            cmd.Transaction = tran;                    
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    tran.Commit();
                }
               catch (Exception exp)
                {
                    string error = exp.Message.ToString();
                    tran.Rollback();
                    MessageBox.Show("Failed to Update. Please check the data");
                }


            //Infragistics.Win.UltraWinGrid.UltraGridBand band = dgDownload.DisplayLayout.Bands[0] as Infragistics.Win.UltraWinGrid.UltraGridBand;
            //band.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
            //foreach (DataColumn dc in DtGrid.Columns)
            //{
            //    if (dc.ColumnName == "mim_slip_reason_cd")
            //    {
            //        band.Columns["mim_slip_reason_cd"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
            //        band.Columns["mim_slip_reason_cd"].ValueList = ddMIMSlipCode;
            //    }
            //    if (dc.ColumnName == "mim_status")
            //    {
            //        band.Columns["mim_status"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
            //        band.Columns["mim_status"].ValueList = ddMIMStatus;
            //    }

            //    if (dc.DataType == System.Type.GetType("System.Int32"))
            //    {
            //        //    band.Columns[dc.ColumnName].MaskInput = "n,nnn";
            //        // band.Columns[dc.ColumnName].UseEditorMaskSettings = true;
            //        band.Columns[dc.ColumnName].MaskInput = "n,nnn";
            //        band.Columns[dc.ColumnName].MaskClipMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals;
            //        band.Columns[dc.ColumnName].MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeBoth;
            //        band.Columns[dc.ColumnName].MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw;
            //        band.Columns[dc.ColumnName].PromptChar = ' ';
            //    }
            //    else if (dc.DataType == System.Type.GetType("System.DateTime"))
            //    {
            //        band.Columns[dc.ColumnName].MaskInput = "mm-dd-yyyy";
            //        //   band.Columns[dc.ColumnName].UseEditorMaskSettings = true;
            //        //    band.Columns[dc.ColumnName].MaskInput = "mm-dd-yyyy";
            //        band.Columns[dc.ColumnName].MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals;
            //    }

                //  lstWorksheet.Items.Add(lst);
                //  this.lstWorksheet.DisplayMember = "Name";
                // this.lstWorksheet.ValueMember = "ID";
                //this.lslWorksheet1.Items.Add(lst);
                // this.lslWorksheet1.DisplayMember = "Name";
                // this.lslWorksheet1.ValueMember = "ID";
                //  }
                //  }

                //for ( int i = 0; i < entries.Count; i++)
                //{
                //    ListEntry entry = entries[i] as ListEntry;

                //    ListEntry.CustomElementCollection elements = entry.Elements;
                //     DataRow myDataRow = table.NewRow();
                //    for (int j = 0; j < elements.Count; j++)
                //    {
                //         myDataRow[j] = elements[j].Value;
                //    }
                //    table.Rows.Add(myDataRow);                       
                //}
                // listView1.Items.Add(item);

                // this.editUriTable.Add(item.Text, entry);
                //export_to_excel();
                //   getExcelcolumnName();
                //  dataGridView1.DataSource = table;

                //if (WorkSheet.Contains("Firm Demand"))
                //    getExcelcolumnName();
                //else
                //{
                //    lblRowCount.Text = "Rows are :" + table.Rows.Count.ToString();
                //    dataGridView1.DataSource = null;
                //    dataGridView1.DataSource = table;
                //    dataGridView1.Columns[dataGridView1.ColumnCount - 1].Width = 200;
                //}
            //}
        }

        private void button1_Click_4(object sender, EventArgs e)
        {
          
          

        }

        private void button7_Click(object sender, EventArgs e)
        {
            loadDynamicGrid();
        }

        private void updateTempToLive()
        {
            string sqlUpdate = " exec sp_app_updateKanbanProd 'REPLENISH' ";// + idString.ToString();
            DataTable dtSave = getdataSet(sqlUpdate).Tables[0];
            if (dtSave != null)
                if (dtSave.Rows.Count > 0)
                    if (dtSave.Rows[0][0].ToString() == "0")
                        MessageBox.Show("Production Replenish has been updated");
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            ClearDyn();
        }


        public void ClearDyn()
        {
            this.textDynID.Text = "";
            this.txtDynMoqQty.Value = 0;            
            //this.txtDynPart.Text = "";
            //this.ddDynSite.Text = "";            
            //this.txtDynSetQty.Value = 0;
           
        }

        private void grdDynamic_AfterSelectChange(object sender, Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs e)
        {
            if (this.grdDynamic.Selected.Rows.Count > 0)
            {
                this.textDynID.Text = grdDynamic.Selected.Rows[0].Cells["MasterID"].Value.ToString();
                this.txtDynMoqQty.Value = grdDynamic.Selected.Rows[0].Cells["TMoqQty"].Value.ToString();
                    //this.txtDynPart.Text = grdDynamic.Selected.Rows[0].Cells["TPartNum"].Value.ToString();
                    //this.ddDynSite.Text = grdDynamic.Selected.Rows[0].Cells["TSiteCode"].Value.ToString();
                    //this.txtDynSetQty.Value = grdDynamic.Selected.Rows[0].Cells["TSetQTY"].Value.ToString();             
            }
        }

        private void txtDynPart_Leave(object sender, EventArgs e)
        {
            //if (this.txtDynPart.Text.Trim() != "")
            //{
            //    if (!validatePart(this.txtDynPart.Text.Trim()))
            //    {
            //        this.txtDynPart.Text = "";
            //        this.txtDynPart.Focus();
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Please enter part num");
            //    this.txtDynPart.Focus();
            //}
        }

        private void btnDynsave_Click(object sender, EventArgs e)
        {
            String sql = " update KanbanMasterTemp set tstatus='N' ";
            int id = 0;

            if (this.textDynID.Text.ToString().Trim() != "")
                id = Convert.ToInt32(this.textDynID.Text.ToString());

             string DeleteFlag = "N";
        //    if (this.chkDynDelete.Checked.ToString().ToUpper() == "TRUE")
         //       DeleteFlag = "Y";

            int MoqQTY = 0;              
            MoqQTY = Convert.ToInt32(this.txtDynMoqQty.Value.ToString());
           
            if (id != 0)
                sql = sql + "  Update KanbanMasterTemp set  TMoqQty='" + MoqQTY +  "', TStatus='N'  where MasterID=" + id;
        

            if (sql != "")
            {
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlTransaction tran = cn.BeginTransaction("Transaction");
                SqlCommand cmd = new SqlCommand(sql, cn);
                cmd.Transaction = tran;
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    tran.Commit();
                    MessageBox.Show("Updated successfully");
                }
                catch (Exception exp)
                {
                    string error = exp.Message.ToString();
                    tran.Rollback();
                    MessageBox.Show("Failed to Update. Please check the data");
                }
                cn.Close();
            }

            loadDynamicGrid();
        }

        private void grdDynamic_ClickCell(object sender, Infragistics.Win.UltraWinGrid.ClickCellEventArgs e)
        {
            if (e.Cell.Column.Index == 19)
            {

                if (e.Cell.Row.Cells["Masterid"].Value.ToString() != "")
                {
                    DeleteRowDataBase(Convert.ToInt32(e.Cell.Row.Cells["Masterid"].Value));
                }

                this.grdDynamic.Rows[e.Cell.Row.Index].Delete();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            String sql = " update KanbanMasterTemp set tstatus='N' ";
            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.grdDynamic.Rows)
            {
                int id = 0;
                if (row.Cells["MasterID"].Text.ToString().Trim() != "")
                    id = Convert.ToInt32(row.Cells["MasterID"].Text.ToString());
               // string site_code = row.Cells["TSiteCode"].Text.ToString();
              //  string part_num = row.Cells["TPartNum"].Text.ToString();
              //  string displayPartNum = row.Cells["TDisplayPN"].Text.ToString();
              //  string OrdType = row.Cells["TOrdType"].Text.ToString();
             //   string OrdPrefix = row.Cells["TOrdPrefix"].Text.ToString();
             //   string ContType = row.Cells["TContType"].Text.ToString();
                string DeleteFlag = "N";
                string OnHold = "N";
                if (row.Cells["TDelete"].Text.ToString().ToUpper() == "TRUE")
                    DeleteFlag = "Y";
                
                if (row.Cells["TOnHold"].Text.ToString().ToUpper() == "TRUE")
                    OnHold = "Y";

                int QtyInCont = 0;
                int QtyOfOrders = 0;
                int MoqQtyOrders = 0;

             //   if (row.Cells["TQtyInCont"].Text.Trim().ToString() != "")

                //    QtyInCont = Convert.ToInt32(row.Cells["TQtyInCont"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
            //    if (row.Cells["TQtyOfOrders"].Text.Trim().ToString() != "")
              //      QtyOfOrders = Convert.ToInt32(row.Cells["TQtyOfOrders"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                if (row.Cells["TMoqQty"].Text.Trim().ToString() != "")
                    MoqQtyOrders = Convert.ToInt32(row.Cells["TMoqQty"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());               
                    if (id != 0)
                   
                    {
                        sql = sql + "  Update KanbanMasterTemp set  TMoqQty='" + MoqQtyOrders + "', TStatus='N'  where MasterID=" + id;
                    }
                
            }
            if (sql != "")
            {

                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlTransaction tran = cn.BeginTransaction("Transaction");
                SqlCommand cmd = new SqlCommand(sql, cn);
                cmd.Transaction = tran;
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    tran.Commit();
                    MessageBox.Show("Updated successfully");
                }

                catch (Exception exp)
                {
                    string error = exp.Message.ToString();
                    tran.Rollback();
                    MessageBox.Show("Failed to Update. Please check the data");
                }
                cn.Close();
            }
            loadGrid();
        }

    

        private void button5_Click(object sender, EventArgs e)
        {
            //connectSharedDoc();
            //create_file_OLEDB();
            //uploadFile();

        }

      

        private void btnLoadDynPart_Click(object sender, EventArgs e)
        {
            string sqlUpdate = " exec sp_LoadDynamicPart  ";// + idString.ToString();
            DataTable dtSave = getdataSet(sqlUpdate).Tables[0];
            if (dtSave != null)
                if (dtSave.Rows.Count > 0)
                    if (dtSave.Rows[0][0].ToString() == "0")
                        MessageBox.Show("Dynamic parts have been loaded");

            loadDynamicGrid();
        }

       

        private void btnUpdateShareDoc_Click(object sender, EventArgs e)
        {
           // connectSharedDoc();
            create_file_OLEDB();
            uploadFile();
        }

        private void btnUpdateProdTable_Click(object sender, EventArgs e)
        {
            string sqlUpdate = " exec sp_app_updateKanbanProd 'REPLENISH' ";// + idString.ToString();
            DataTable dtSave = getdataSet(sqlUpdate).Tables[0];
            if (dtSave != null)
                if (dtSave.Rows.Count > 0)
                    if (dtSave.Rows[0][0].ToString() == "0")
                        MessageBox.Show("Production Replenish has been updated");
        }

        private void btnUpdateMenloMOQ_Click(object sender, EventArgs e)
        {
            ImportsharedocToExcel();
           // connectSharedDoc();
        }

        private void gridKanban_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {

        }

        private void button1_Click_5(object sender, EventArgs e)
        {
            this.ddFilterSite.Text="";
            this.comboPartType.Text = "";
            this.txtPartNum.Text = "";
            this.ultraTextEditorDescription.Text = "";
        }

        private void ddFilterRepSite_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {

        }

        private void ddFilterRepSite_RowSelected(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
        {
            if (ddFilterRepSite.Text=="NL" || ddFilterRepSite.Text=="ALL")
                btnUpdateMenloMOQ.Visible=false;
            else
                 btnUpdateMenloMOQ.Visible=true;
        }

    }
}
