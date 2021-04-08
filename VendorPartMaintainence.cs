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
///using customcontrol;

using System.Globalization;

//using Google.GData.Client;
//using Google.GData.Extensions;
//using Google.GData.Spreadsheets;

//using System.Diagnostics;
//using DotNetOpenAuth.OAuth2;
//using Google.Apis.Authentication.OAuth2;
//using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
//using Google.Apis.Helper;
//using Google.Apis.Authentication;
//using Google.Apis.Tasks.v1;
//using Google.Apis.Tasks.v1.Data;
//using Google.Apis.Drive.v2.Data;
//using Google.Apis.Drive.v2;
//using Google.Apis.Util;



namespace Version3
{
    public partial class VendorCommit : Form
    {
        //Google.GData.Documents.DocumentsService docService;
        //Google.GData.Spreadsheets.SpreadsheetsService spreadsheetService;
        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        //Boolean loaded;
        public VendorCommit()
        {
            InitializeComponent();

          // 
            //loadVCVendor();
            //loadSite();
            loadStatus();
            // getWorksheet();        
        }

        //System.Collections.Hashtable editUriTable;
        //SpreadsheetsService service;
        //string fileName = "MIM/Menlo Demand Tracker Test";
        //string fileName = "Vendor Commit Sheet Template";
        //string worksheetname;
        //AtomLink listLink;
        //AtomLink cellsLink;
        //AtomEntryCollection wEntries;
        //System.Collections.Hashtable editUriTableCell;


        //int wk_num_04;
        //string date_04;
        //int year_04;

        //int wk_num_03;
        //string date_03;
        //int year_03;

        //int wk_num_02;
        //string date_02;
        //int year_02;

        //int wk_num_01;
        //string date_01;
        //int year_01;

        //int wk_num_0;
        //string date_0;
        //int year_0;

        //int wk_num_1;
        //string date_1;
        //int year_1;

        //int wk_num_2;
        //string date_2;
        //int year_2;

        //int wk_num_3;
        //string date_3;
        //int year_3;

        //int wk_num_4;
        //string date_4;
        //int year_4;

        //int wk_num_5;
        //string date_5;

        //int wk_num_6;
        //string date_6;

        //int wk_num_7;
        //string date_7;

        //int wk_num_8;
        //string date_8;

        //int wk_num_9;
        //string date_9;

        //int wk_num_10;
        //string date_10;

        //int wk_num_11;
        //string date_11;

        //int wk_num_12;
        //string date_12;

        //int wk_num_13;
        //string date_13;

        //int wk_num_14;
        //string date_14;

        //int wk_num_15;
        //string date_15;

        //int wk_num_16;
        //string date_16;

        //int wk_num_17;
        //string date_17;

        //int wk_num_18;
        //string date_18;

        //int wk_num_19;
        //string date_19;

        //int wk_num_20;
        //string date_20;

        //int wk_num_21;
        //string date_21;

        //int wk_num_22;
        //string date_22;

        //int wk_num_23;
        //string date_23;

        //int wk_num_24;
        //string date_24;

        //int wk_num_25;
        //string date_25;
        
        //int year_5;
        //int year_6;
        //int year_7;
        //int year_8;
        //int year_9;
        //int year_10;
        //int year_11;
        //int year_12;
        //int year_13;
        //int year_14;
        //int year_15;
        //int year_16;
        //int year_17;
        //int year_18;
        //int year_19;
        //int year_20;
        //int year_21;
        //int year_22;
        //int year_23;
        //int year_24;
        //int year_25;


        //public void connectSharedDoc()
        //{
        //    service = new SpreadsheetsService("service");
        //    service.setUserCredentials("tedatmim@gmail.com", "mimi100");
        //    editUriTable = new System.Collections.Hashtable();
        //    SpreadsheetQuery qrySpreadsheet = new SpreadsheetQuery();
        //    SpreadsheetFeed sFeed = service.Query(qrySpreadsheet);
        //    bool bFound = false;

        //    AtomEntryCollection entries = sFeed.Entries;
        //    for (int i = 0; i < entries.Count; i++)
        //    {
        //        if (entries[i].Title.Text == fileName)
        //        {
        //            // Get the worksheets feed URI
        //            AtomLink worksheetsLink = entries[i].Links.FindService(GDataSpreadsheetsNameTable.WorksheetRel, AtomLink.ATOM_TYPE);
        //            WorksheetQuery query = new WorksheetQuery(worksheetsLink.HRef.Content);
        //            WorksheetFeed wFeed = service.Query(query);
        //            wEntries = wFeed.Entries;
        //            cellsLink = wEntries[0].Links.FindService(GDataSpreadsheetsNameTable.CellRel, AtomLink.ATOM_TYPE);
        //            listLink = wEntries[0].Links.FindService(GDataSpreadsheetsNameTable.ListRel, AtomLink.ATOM_TYPE);
        //            /* 
        //             getCellFeed(cellsLink.HRef.Content);
        //             update_cell(1, 1, " Please wait. Updating..............");
        //             for (uint r = 6; r <= ((WorksheetEntry)wEntries[0]).RowCount.Count; r++)
        //                 for (uint c = 1; c <= ((WorksheetEntry)wEntries[0]).ColCount.Count; c++)
        //                 {
        //                     if (c != 18)
        //                     {
        //                         update_cell(r, c, "");
        //                     }
        //                 }
        //             */
        //            break;
        //        }
        //    }
        //}

        //public void loadVendor()
        //{
        //    string strsqlVendor = " select vend_id,ep_vend_code as [Vendor] from dbo.vc_vend_master where active='Y' order by ep_vend_code";

        //    DataSet dsVendor = getdataSet(strsqlVendor);
        //    DataRow dr = dsVendor.Tables[0].NewRow();
        //    dr["vend_id"] = "0";
        //    dr["Vendor"] = "Select";
        //    dsVendor.Tables[0].Rows.InsertAt(dr, 0);
        //    cmbVendor.DataSource = dsVendor.Tables[0];
        //    cmbVendor.DisplayMember = "Vendor";
        //    cmbVendor.ValueMember = "vend_id";

        //    //cmbVendPart.DataSource = dsVendor.Tables[0];
        //    //cmbVendPart.DisplayMember = "Vendor";
        //    //cmbVendPart.ValueMember = "vend_id";

        //    //DataTable dt = new DataTable();
        //    //DataColumn display=new DataColumn();
        //    //DataColumn value=new DataColumn();

        //    //dt.Columns.Add(display);
        //    //    dt.Columns.Add(value);
        //    //DataRow dr=dt.NewRow();
        //    //dr("Active",
        //    //dt.Rows.Add(new DataRow({"Active","ACTIVE"}));
        //    //cmbStatus.DataSource
        //}

        //public void loadVCVendor()
        //{
        //    string strsqlVendor = " select vend_id,ep_vend_code as [Vendor] from dbo.vc_vend_master where active='Y' order by ep_vend_code";

        //    DataSet dsVendor = getdataSet(strsqlVendor);
        //    DataRow dr = dsVendor.Tables[0].NewRow();
        //    dr["vend_id"] = "0";
        //    dr["Vendor"] = "Select";
        //    dsVendor.Tables[0].Rows.InsertAt(dr, 0);
        //    this.cmbVCVendor.DataSource = dsVendor.Tables[0];
        //    cmbVCVendor.DisplayMember = "Vendor";
        //    cmbVCVendor.ValueMember = "vend_id";

        //    //cmbVendPart.DataSource = dsVendor.Tables[0];
        //    //cmbVendPart.DisplayMember = "Vendor";
        //    //cmbVendPart.ValueMember = "vend_id";

        //    //DataTable dt = new DataTable();
        //    //DataColumn display=new DataColumn();
        //    //DataColumn value=new DataColumn();

        //    //dt.Columns.Add(display);
        //    //    dt.Columns.Add(value);
        //    //DataRow dr=dt.NewRow();
        //    //dr("Active",
        //    //dt.Rows.Add(new DataRow({"Active","ACTIVE"}));


        //    //cmbStatus.DataSource

        //}

        //public void clearFileCreateAttribute()
        //{
        //    this.lblstatusAT.Text = "";
        //    this.lblstatusNL.Text = "";
        //    this.updateDateAT.Text = "";
        //    this.updateDateNL.Text = "";
        //}
        //public void loadVendorFileCreate()
        //{
        //    string strsqlVendor = " select vend_id,ep_vend_code as [Vendor] from dbo.vc_vend_master where active='Y' order by ep_vend_code";

        //    DataSet dsVendor = getdataSet(strsqlVendor);
        //    DataRow dr = dsVendor.Tables[0].NewRow();
        //    dr["vend_id"] = "0";
        //    dr["Vendor"] = "Select";
        //    dsVendor.Tables[0].Rows.InsertAt(dr, 0);
        //    this.cmbVendorFileCreate.DataSource = dsVendor.Tables[0];
        //    cmbVendorFileCreate.DisplayMember = "Vendor";
        //    cmbVendorFileCreate.ValueMember = "vend_id";

        //}

        public void loadComboVendor()
        {
            string strsqlVendor = " select vend_id,ep_vend_code as [Vendor] from dbo.vc_vend_master where active='Y'";

            DataSet dsVendor = getdataSet(strsqlVendor);
            DataRow dr = dsVendor.Tables[0].NewRow();
            dr["vend_id"] = "0";
            dr["Vendor"] = "Select";
            dsVendor.Tables[0].Rows.InsertAt(dr, 0);
            //cmbVendor.DataSource = dsVendor.Tables[0];
            //cmbVendor.DisplayMember = "Vendor";
            //cmbVendor.ValueMember = "vend_id";
            cmbVendPart.DataSource = dsVendor.Tables[0];
            cmbVendPart.DisplayMember = "Vendor";
            cmbVendPart.ValueMember = "vend_id";
        }


        public DataSet getdataSet(string sql)
        {
            SqlConnection cn = new SqlConnection(constr);
            try
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand(sql, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                cmd.CommandTimeout = 0;
                DataSet ds = new DataSet();
                da.Fill(ds);
                cn.Close();
                da.Dispose();
                cmd.Dispose();
                return ds;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex.Message.ToString());
            }
            finally
            {
                if (cn.State == ConnectionState.Open)
                {
                    cn.Close();
                }
            }
            return null;
        }

        //public void loadSite()
        //{
        //    string strsqlSite = " select distinct bundesland as site_code,bundesland as Site_desc from locations where void<>'V' and  bundesland<>'' ";
        //    DataSet dsSite = getdataSet(strsqlSite);
        //    DataRow dr = dsSite.Tables[0].NewRow();
        //    dr["site_code"] = "0";
        //    dr["Site_desc"] = "Select";
        //    dsSite.Tables[0].Rows.InsertAt(dr, 0);
        //    cmbSite.DataSource = dsSite.Tables[0];
        //    cmbSite.DisplayMember = "Site_desc";
        //    cmbSite.ValueMember = "site_code";
        //}

        public void loadStatus()
        {
            string strsqlActive = " select distinct case when Active='Y' then 'ACTIVE' else 'INACTIVE' end as disp_val, Active as value from vc_vend_part  ";
            DataSet dsActive = getdataSet(strsqlActive);
            DataRow dr = dsActive.Tables[0].NewRow();
            dr["disp_val"] = "ALL";
            dr["value"] = "0";
            dsActive.Tables[0].Rows.InsertAt(dr, 0);
            this.cmbStatus.DataSource = dsActive.Tables[0];
            cmbStatus.SelectedIndex = 1;
            cmbStatus.DisplayMember = "disp_val";
            cmbStatus.ValueMember = "value";
        }

        DataTable distinctValues;

        DataSet dsPart;
        //public void getPart()
        //{

        //    string strsqlPart = " exec m_vendor_demand_supply  " + cmbVendor.SelectedValue + ",'" + this.cmbSite.SelectedValue + "'";/// select part_num,short_desc from vc_vend_part where vend_id=" + cmbVendor.SelectedValue;


        //    //try
        //    //  {
        //    dsPart = getdataSet(strsqlPart);
        //    DataSet dsWeek = getdataSet("select top 26  convert(varchar,date,101) as date,mweek,myear from m_week_number_table where date >=(getdate()-42) ");
        //    DataSet dsCommitReq = getdataSet("select *,convert(varchar,create_date,101) as created_date from vc_commit where (status='N' or status='P') and vend_id='" + cmbVendor.SelectedValue + "' and site='" + this.cmbSite.SelectedValue + "' and date>getdate()-42 order by year,week");
        //   // DataSet dsOffset = getdataSet("select *,dbo.udf_GetISOWeekNumberFromDate(dateadd(day, vend.lead_time_air,wk.date)) air_time_week,dbo.udf_GetISOWeekNumberFromDate(dateadd(day, isnull(vend.lead_time_ocean,0),wk.date)) as ocean_time_week,dbo.udf_GetISOWeekNumberFromDate(dateadd(day, vend.lead_time_ground,wk.date)) as ground_time_week from (select part_num,lead_time_air,lead_time_ocean,lead_time_ground from vc_vend_part   where vend_id=" + this.cmbVendor.SelectedValue + " ) as vend cross join (select top 26  convert(varchar,date,101) as date,mweek,myear from m_week_number_table where date >=(getdate()-49)) as wk");

        //     DataSet dsOffset = getdataSet("select part_num,lead_time_air,lead_time_ocean, isnull(lead_time_ground,0) as lead_time_ground,date,mweek,myear,dbo.udf_GetISOWeekNumberFromDate(dateadd(day, vend.lead_time_air,wk.date)) air_time_week,isnull(dbo.udf_GetISOWeekNumberFromDate(dateadd(day, vend.lead_time_ocean,wk.date)),0) as ocean_time_week,dbo.udf_GetISOWeekNumberFromDate(dateadd(day, isnull(vend.lead_time_ground,0),wk.date)) as ground_time_week from (select part_num,lead_time_air,lead_time_ocean,lead_time_ground from vc_vend_part   where vend_id=" + this.cmbVendor.SelectedValue + " ) as vend cross join (select top 26  convert(varchar,date,101) as date,mweek,myear from m_week_number_table where date >=(getdate()-49)) as wk");
        //    lblCommitId.Text = "";
        //    customcontrol.UserControl2 ctrlHeader = new UserControl2();
        //    ctrlHeader.Name = "header";
        //    ctrlHeader.Width = 1515;
        //    if (dsCommitReq.Tables.Count > 0)
        //    {
        //        if (dsCommitReq.Tables[0].Rows.Count > 0)
        //        {
        //            lblCreateDate.Text = dsCommitReq.Tables[0].Rows[0]["created_date"].ToString();
        //            if (dsCommitReq.Tables[0].Rows[0]["status"].ToString() == "N")
        //            {
        //                lblStatus.Text = "NEW";
        //                button1.Enabled = true;
        //                button1.BackColor = System.Drawing.Color.White;
        //                button2.Enabled = false;
        //                button2.BackColor = System.Drawing.Color.IndianRed;
        //                button3.Enabled = true;
        //                button3.BackColor = System.Drawing.Color.White;
        //            }
        //            else if (dsCommitReq.Tables[0].Rows[0]["status"].ToString() == "P")
        //            {
        //                lblStatus.Text = "READY To SEND";
        //                button1.Enabled = false;
        //                button1.BackColor = System.Drawing.Color.IndianRed;
        //                button2.Enabled = true;
        //                button2.BackColor = System.Drawing.Color.White;
        //                button3.Enabled = false;
        //                button3.BackColor = System.Drawing.Color.IndianRed;
        //            }
        //        }
        //        else
        //        {
        //            lblStatus.Text = "";
        //            button1.Enabled = true;
        //            button2.Enabled = false;
        //            button3.Enabled = false;
        //            button1.BackColor = System.Drawing.Color.White;
        //            button2.BackColor = System.Drawing.Color.IndianRed;
        //            button3.BackColor = System.Drawing.Color.IndianRed;
        //        }
        //    }

        //    ctrlHeader.v_wk_04 = Convert.ToInt32(dsWeek.Tables[0].Rows[0]["mweek"].ToString());
        //    ctrlHeader.v_yr_04 = Convert.ToInt32(dsWeek.Tables[0].Rows[0]["myear"].ToString());
        //    ctrlHeader.v_date_04 = dsWeek.Tables[0].Rows[0]["date"].ToString();
        //    wk_num_04 = Convert.ToInt32(dsWeek.Tables[0].Rows[0]["mweek"].ToString());
        //    date_04 = dsWeek.Tables[0].Rows[0]["date"].ToString();
        //    year_04 = Convert.ToInt32(dsWeek.Tables[0].Rows[0]["myear"].ToString());

        //    ctrlHeader.v_wk_03 = Convert.ToInt32(dsWeek.Tables[0].Rows[1]["mweek"].ToString());
        //    ctrlHeader.v_yr_03 = Convert.ToInt32(dsWeek.Tables[0].Rows[1]["myear"].ToString());
        //    ctrlHeader.v_date_03 = dsWeek.Tables[0].Rows[1]["date"].ToString();
        //    wk_num_03 = Convert.ToInt32(dsWeek.Tables[0].Rows[1]["mweek"].ToString());
        //    date_03 = dsWeek.Tables[0].Rows[1]["date"].ToString();
        //    year_03 = Convert.ToInt32(dsWeek.Tables[0].Rows[1]["myear"].ToString());

        //    ctrlHeader.v_wk_02 = Convert.ToInt32(dsWeek.Tables[0].Rows[2]["mweek"].ToString());
        //    ctrlHeader.v_yr_02 = Convert.ToInt32(dsWeek.Tables[0].Rows[2]["myear"].ToString());
        //    ctrlHeader.v_date_02 = dsWeek.Tables[0].Rows[2]["date"].ToString();
        //    wk_num_02 = Convert.ToInt32(dsWeek.Tables[0].Rows[2]["mweek"].ToString());
        //    date_02 = dsWeek.Tables[0].Rows[2]["date"].ToString();
        //    year_02 = Convert.ToInt32(dsWeek.Tables[0].Rows[2]["myear"].ToString());

        //    ctrlHeader.v_wk_01 = Convert.ToInt32(dsWeek.Tables[0].Rows[3]["mweek"].ToString());
        //    ctrlHeader.v_yr_01 = Convert.ToInt32(dsWeek.Tables[0].Rows[3]["myear"].ToString());
        //    ctrlHeader.v_date_01 = dsWeek.Tables[0].Rows[3]["date"].ToString();
        //    wk_num_01 = Convert.ToInt32(dsWeek.Tables[0].Rows[3]["mweek"].ToString());
        //    date_01 = dsWeek.Tables[0].Rows[3]["date"].ToString();
        //    year_01 = Convert.ToInt32(dsWeek.Tables[0].Rows[3]["myear"].ToString());

        //    ctrlHeader.v_wk_0 = Convert.ToInt32(dsWeek.Tables[0].Rows[4]["mweek"].ToString());
        //    ctrlHeader.v_yr_0 = Convert.ToInt32(dsWeek.Tables[0].Rows[4]["myear"].ToString());
        //    ctrlHeader.v_date_0 = dsWeek.Tables[0].Rows[4]["date"].ToString();
        //    wk_num_0 = Convert.ToInt32(dsWeek.Tables[0].Rows[4]["mweek"].ToString());
        //    date_0 = dsWeek.Tables[0].Rows[4]["date"].ToString();
        //    year_0 = Convert.ToInt32(dsWeek.Tables[0].Rows[4]["myear"].ToString());

        //    ctrlHeader.v_wk_1 = Convert.ToInt32(dsWeek.Tables[0].Rows[5]["mweek"].ToString());
        //    ctrlHeader.v_yr_1 = Convert.ToInt32(dsWeek.Tables[0].Rows[5]["myear"].ToString());
        //    ctrlHeader.v_date_1 = dsWeek.Tables[0].Rows[5]["date"].ToString();
        //    wk_num_1 = Convert.ToInt32(dsWeek.Tables[0].Rows[5]["mweek"].ToString());
        //    date_1 = dsWeek.Tables[0].Rows[5]["date"].ToString();
        //    year_1 = Convert.ToInt32(dsWeek.Tables[0].Rows[5]["myear"].ToString());

        //    ctrlHeader.v_wk_2 = Convert.ToInt32(dsWeek.Tables[0].Rows[6]["mweek"].ToString());
        //    ctrlHeader.v_yr_2 = Convert.ToInt32(dsWeek.Tables[0].Rows[6]["myear"].ToString());
        //    ctrlHeader.v_date_2 = dsWeek.Tables[0].Rows[6]["date"].ToString();
        //    wk_num_2 = Convert.ToInt32(dsWeek.Tables[0].Rows[6]["mweek"].ToString());
        //    date_2 = dsWeek.Tables[0].Rows[6]["date"].ToString();
        //    year_2 = Convert.ToInt32(dsWeek.Tables[0].Rows[6]["myear"].ToString());


        //    ctrlHeader.v_wk_3 = Convert.ToInt32(dsWeek.Tables[0].Rows[7]["mweek"].ToString());
        //    ctrlHeader.v_yr_3 = Convert.ToInt32(dsWeek.Tables[0].Rows[7]["myear"].ToString());
        //    ctrlHeader.v_date_3 = dsWeek.Tables[0].Rows[7]["date"].ToString();
        //    wk_num_3 = Convert.ToInt32(dsWeek.Tables[0].Rows[7]["mweek"].ToString());
        //    date_3 = dsWeek.Tables[0].Rows[7]["date"].ToString();
        //    year_3 = Convert.ToInt32(dsWeek.Tables[0].Rows[7]["myear"].ToString());

        //    ctrlHeader.v_wk_4 = Convert.ToInt32(dsWeek.Tables[0].Rows[8]["mweek"].ToString());
        //    ctrlHeader.v_yr_4 = Convert.ToInt32(dsWeek.Tables[0].Rows[8]["myear"].ToString());
        //    ctrlHeader.v_date_4 = dsWeek.Tables[0].Rows[8]["date"].ToString();
        //    wk_num_4 = Convert.ToInt32(dsWeek.Tables[0].Rows[8]["mweek"].ToString());
        //    date_4 = dsWeek.Tables[0].Rows[8]["date"].ToString();
        //    year_4 = Convert.ToInt32(dsWeek.Tables[0].Rows[8]["myear"].ToString());

        //    ctrlHeader.v_wk_5 = Convert.ToInt32(dsWeek.Tables[0].Rows[9]["mweek"].ToString());
        //    ctrlHeader.v_yr_5 = Convert.ToInt32(dsWeek.Tables[0].Rows[9]["myear"].ToString());
        //    ctrlHeader.v_date_5 = dsWeek.Tables[0].Rows[9]["date"].ToString();
        //    wk_num_5 = Convert.ToInt32(dsWeek.Tables[0].Rows[9]["mweek"].ToString());
        //    date_5 = dsWeek.Tables[0].Rows[9]["date"].ToString();
        //    year_5 = Convert.ToInt32(dsWeek.Tables[0].Rows[9]["myear"].ToString());

        //    ctrlHeader.v_wk_6 = Convert.ToInt32(dsWeek.Tables[0].Rows[10]["mweek"].ToString());
        //    ctrlHeader.v_yr_6 = Convert.ToInt32(dsWeek.Tables[0].Rows[10]["myear"].ToString());
        //    ctrlHeader.v_date_6 = dsWeek.Tables[0].Rows[10]["date"].ToString();
        //    wk_num_6 = Convert.ToInt32(dsWeek.Tables[0].Rows[10]["mweek"].ToString());
        //    date_6 = dsWeek.Tables[0].Rows[10]["date"].ToString();
        //    year_6 = Convert.ToInt32(dsWeek.Tables[0].Rows[10]["myear"].ToString());

        //    ctrlHeader.v_wk_7 = Convert.ToInt32(dsWeek.Tables[0].Rows[11]["mweek"].ToString());
        //    ctrlHeader.v_yr_7 = Convert.ToInt32(dsWeek.Tables[0].Rows[11]["myear"].ToString());
        //    ctrlHeader.v_date_7 = dsWeek.Tables[0].Rows[11]["date"].ToString();
        //    wk_num_7 = Convert.ToInt32(dsWeek.Tables[0].Rows[11]["mweek"].ToString());
        //    date_7 = dsWeek.Tables[0].Rows[11]["date"].ToString();
        //    year_7 = Convert.ToInt32(dsWeek.Tables[0].Rows[11]["myear"].ToString());

        //    ctrlHeader.v_wk_8 = Convert.ToInt32(dsWeek.Tables[0].Rows[12]["mweek"].ToString());
        //    ctrlHeader.v_yr_8 = Convert.ToInt32(dsWeek.Tables[0].Rows[12]["myear"].ToString());
        //    ctrlHeader.v_date_8 = dsWeek.Tables[0].Rows[12]["date"].ToString();
        //    wk_num_8 = Convert.ToInt32(dsWeek.Tables[0].Rows[12]["mweek"].ToString());
        //    date_8 = dsWeek.Tables[0].Rows[12]["date"].ToString();
        //    year_8 = Convert.ToInt32(dsWeek.Tables[0].Rows[12]["myear"].ToString());

        //    ctrlHeader.v_wk_9 = Convert.ToInt32(dsWeek.Tables[0].Rows[13]["mweek"].ToString());
        //    ctrlHeader.v_yr_9 = Convert.ToInt32(dsWeek.Tables[0].Rows[13]["myear"].ToString());
        //    ctrlHeader.v_date_9 = dsWeek.Tables[0].Rows[13]["date"].ToString();
        //    wk_num_9 = Convert.ToInt32(dsWeek.Tables[0].Rows[13]["mweek"].ToString());
        //    date_9 = dsWeek.Tables[0].Rows[13]["date"].ToString();
        //    year_9 = Convert.ToInt32(dsWeek.Tables[0].Rows[13]["myear"].ToString());

        //    ctrlHeader.v_wk_10 = Convert.ToInt32(dsWeek.Tables[0].Rows[14]["mweek"].ToString());
        //    ctrlHeader.v_yr_10 = Convert.ToInt32(dsWeek.Tables[0].Rows[14]["myear"].ToString());
        //    ctrlHeader.v_date_10 = dsWeek.Tables[0].Rows[14]["date"].ToString();
        //    wk_num_10 = Convert.ToInt32(dsWeek.Tables[0].Rows[14]["mweek"].ToString());
        //    date_10 = dsWeek.Tables[0].Rows[14]["date"].ToString();
        //    year_10 = Convert.ToInt32(dsWeek.Tables[0].Rows[14]["myear"].ToString());

        //    ctrlHeader.v_wk_11 = Convert.ToInt32(dsWeek.Tables[0].Rows[15]["mweek"].ToString());
        //    ctrlHeader.v_yr_11 = Convert.ToInt32(dsWeek.Tables[0].Rows[15]["myear"].ToString());
        //    ctrlHeader.v_date_11 = dsWeek.Tables[0].Rows[15]["date"].ToString();
        //    wk_num_11 = Convert.ToInt32(dsWeek.Tables[0].Rows[15]["mweek"].ToString());
        //    date_11 = dsWeek.Tables[0].Rows[15]["date"].ToString();
        //    year_11 = Convert.ToInt32(dsWeek.Tables[0].Rows[15]["myear"].ToString());

        //    ctrlHeader.v_wk_12 = Convert.ToInt32(dsWeek.Tables[0].Rows[16]["mweek"].ToString());
        //    ctrlHeader.v_yr_12 = Convert.ToInt32(dsWeek.Tables[0].Rows[16]["myear"].ToString());
        //    ctrlHeader.v_date_12 = dsWeek.Tables[0].Rows[16]["date"].ToString();
        //    wk_num_12 = Convert.ToInt32(dsWeek.Tables[0].Rows[16]["mweek"].ToString());
        //    date_12 = dsWeek.Tables[0].Rows[16]["date"].ToString();
        //    year_12 = Convert.ToInt32(dsWeek.Tables[0].Rows[16]["myear"].ToString());

        //    ctrlHeader.v_wk_13 = Convert.ToInt32(dsWeek.Tables[0].Rows[17]["mweek"].ToString());
        //    ctrlHeader.v_yr_13 = Convert.ToInt32(dsWeek.Tables[0].Rows[17]["myear"].ToString());
        //    ctrlHeader.v_date_13 = dsWeek.Tables[0].Rows[17]["date"].ToString();
        //    wk_num_13 = Convert.ToInt32(dsWeek.Tables[0].Rows[17]["mweek"].ToString());
        //    date_13 = dsWeek.Tables[0].Rows[17]["date"].ToString();
        //    year_13 = Convert.ToInt32(dsWeek.Tables[0].Rows[17]["myear"].ToString());

        //    ctrlHeader.v_wk_14 = Convert.ToInt32(dsWeek.Tables[0].Rows[18]["mweek"].ToString());
        //    ctrlHeader.v_yr_14 = Convert.ToInt32(dsWeek.Tables[0].Rows[18]["myear"].ToString());
        //    ctrlHeader.v_date_14 = dsWeek.Tables[0].Rows[18]["date"].ToString();
        //    wk_num_14 = Convert.ToInt32(dsWeek.Tables[0].Rows[18]["mweek"].ToString());
        //    date_14 = dsWeek.Tables[0].Rows[18]["date"].ToString();
        //    year_14 = Convert.ToInt32(dsWeek.Tables[0].Rows[18]["myear"].ToString());

        //    ctrlHeader.v_wk_15 = Convert.ToInt32(dsWeek.Tables[0].Rows[19]["mweek"].ToString());
        //    ctrlHeader.v_yr_15 = Convert.ToInt32(dsWeek.Tables[0].Rows[19]["myear"].ToString());
        //    ctrlHeader.v_date_15 = dsWeek.Tables[0].Rows[19]["date"].ToString();
        //    wk_num_15 = Convert.ToInt32(dsWeek.Tables[0].Rows[19]["mweek"].ToString());
        //    date_15 = dsWeek.Tables[0].Rows[19]["date"].ToString();
        //    year_15 = Convert.ToInt32(dsWeek.Tables[0].Rows[19]["myear"].ToString());

        //    ctrlHeader.v_wk_16 = Convert.ToInt32(dsWeek.Tables[0].Rows[20]["mweek"].ToString());
        //    ctrlHeader.v_yr_16 = Convert.ToInt32(dsWeek.Tables[0].Rows[20]["myear"].ToString());
        //    ctrlHeader.v_date_16 = dsWeek.Tables[0].Rows[20]["date"].ToString();
        //    wk_num_16 = Convert.ToInt32(dsWeek.Tables[0].Rows[20]["mweek"].ToString());
        //    date_16 = dsWeek.Tables[0].Rows[20]["date"].ToString();
        //    year_16 = Convert.ToInt32(dsWeek.Tables[0].Rows[20]["myear"].ToString());


        //    ctrlHeader.v_wk_17 = Convert.ToInt32(dsWeek.Tables[0].Rows[21]["mweek"].ToString());
        //    ctrlHeader.v_yr_17 = Convert.ToInt32(dsWeek.Tables[0].Rows[21]["myear"].ToString());
        //    ctrlHeader.v_date_17 = dsWeek.Tables[0].Rows[21]["date"].ToString();
        //    wk_num_17 = Convert.ToInt32(dsWeek.Tables[0].Rows[21]["mweek"].ToString());
        //    date_17 = dsWeek.Tables[0].Rows[21]["date"].ToString();
        //    year_17 = Convert.ToInt32(dsWeek.Tables[0].Rows[21]["myear"].ToString());

        //    ctrlHeader.v_wk_18 = Convert.ToInt32(dsWeek.Tables[0].Rows[22]["mweek"].ToString());
        //    ctrlHeader.v_yr_18 = Convert.ToInt32(dsWeek.Tables[0].Rows[22]["myear"].ToString());
        //    ctrlHeader.v_date_18 = dsWeek.Tables[0].Rows[22]["date"].ToString();
        //    wk_num_18 = Convert.ToInt32(dsWeek.Tables[0].Rows[22]["mweek"].ToString());
        //    date_18 = dsWeek.Tables[0].Rows[22]["date"].ToString();
        //    year_18 = Convert.ToInt32(dsWeek.Tables[0].Rows[22]["myear"].ToString());

        //    ctrlHeader.v_wk_19 = Convert.ToInt32(dsWeek.Tables[0].Rows[23]["mweek"].ToString());
        //    ctrlHeader.v_yr_19 = Convert.ToInt32(dsWeek.Tables[0].Rows[23]["myear"].ToString());
        //    ctrlHeader.v_date_19 = dsWeek.Tables[0].Rows[23]["date"].ToString();
        //    wk_num_19 = Convert.ToInt32(dsWeek.Tables[0].Rows[23]["mweek"].ToString());
        //    date_19 = dsWeek.Tables[0].Rows[23]["date"].ToString();
        //    year_19 = Convert.ToInt32(dsWeek.Tables[0].Rows[23]["myear"].ToString());

        //    ctrlHeader.v_wk_20 = Convert.ToInt32(dsWeek.Tables[0].Rows[24]["mweek"].ToString());
        //    ctrlHeader.v_yr_20 = Convert.ToInt32(dsWeek.Tables[0].Rows[24]["myear"].ToString());
        //    ctrlHeader.v_date_20 = dsWeek.Tables[0].Rows[24]["date"].ToString();
        //    wk_num_20 = Convert.ToInt32(dsWeek.Tables[0].Rows[24]["mweek"].ToString());
        //    date_20 = dsWeek.Tables[0].Rows[24]["date"].ToString();
        //    year_20 = Convert.ToInt32(dsWeek.Tables[0].Rows[24]["myear"].ToString());



        //    TableLayoutPanel tableLayoutPanelheader = new TableLayoutPanel();

        //    tableLayoutPanelheader.Width = pnlHeader.Width;
        //    tableLayoutPanelheader.Height = pnlHeader.Height;

        //    tableLayoutPanelheader.Font = new Font("Verdana", 8, FontStyle.Regular);
        //    //tableLayoutPanel1.CellBorderStyle = CellBorderStyle.FixedSingle;
        //    tableLayoutPanelheader.AutoScroll = true;
        //    tableLayoutPanelheader.ColumnCount = 5;
        //    for (int i = 0; i < tableLayoutPanelheader.ColumnCount; i++)
        //    {
        //        if (i == 0)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 150);
        //            tableLayoutPanelheader.ColumnStyles.Add(cs);
        //        }
        //        //if (i == 1)
        //        //{
        //        //    ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 60);
        //        //    tableLayoutPanelheader.ColumnStyles.Add(cs);
        //        //}
        //        if (i == 1)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 60);
        //            tableLayoutPanelheader.ColumnStyles.Add(cs);
        //        }
        //        if (i == 2)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 60);
        //            tableLayoutPanelheader.ColumnStyles.Add(cs);
        //        }
        //        if (i == 3)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 60);
        //            tableLayoutPanelheader.ColumnStyles.Add(cs);
        //        }

        //        if (i == 4)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 1515);
        //            tableLayoutPanelheader.ColumnStyles.Add(cs);
        //        }

        //    }
        //    /*if (dsPart != null)
        //    {
        //        //dataGridView1.DataSource = dsPart.Tables[0];
        //        for (int j = 0; j < panel2.Controls.Count; j++)
        //        {
        //            panel2.Controls.RemoveAt(j);
        //        }*/

        //    //tableLayoutPanelheader.RowCount = dsPart.Tables[0].Rows.Count+1;
        //    Label lblPart = new Label();
        //    lblPart.Text = "PART NUM";
        //    lblPart.Width = 130;
        //    lblPart.TextAlign = ContentAlignment.MiddleLeft;

        //    lblPart.Font = new Font(this.Font, FontStyle.Bold);

        //    //Label lblPartDes = new Label();
        //    //lblPartDes.Text = " DESC";
        //    //lblPartDes.Font = new Font(this.Font, FontStyle.Bold);
        //    //lblPartDes.Width =60;
        //    //lblPartDes.TextAlign = ContentAlignment.MiddleCenter;
        //    Label lblStock_oh = new Label();
        //    lblStock_oh.Text = "OH";
        //    lblStock_oh.Font = new Font(this.Font, FontStyle.Bold);
        //    lblStock_oh.Width = 70;
        //    lblStock_oh.TextAlign = ContentAlignment.MiddleCenter;

        //    Label lblStock_intransit = new Label();
        //    lblStock_intransit.Text = "TRN";
        //    lblStock_intransit.Font = new Font(this.Font, FontStyle.Bold);
        //    lblStock_intransit.Width = 70;
        //    lblStock_intransit.TextAlign = ContentAlignment.MiddleCenter;

        //    Label lblStock = new Label();
        //    lblStock.Text = "STK";
        //    lblStock.Font = new Font(this.Font, FontStyle.Bold);
        //    lblStock.Width = 70;
        //    lblStock.TextAlign = ContentAlignment.MiddleCenter;

        //    Label lblWeekNum = new Label();
        //    lblWeekNum.Text = "";

        //    //  lblWeekNum.Font = new Font(this.Font, FontStyle.Bold);

        //    tableLayoutPanelheader.Controls.Add(lblPart, 0, 0);
        //   // tableLayoutPanelheader.Controls.Add(lblPartDes, 1, 0);
        //    tableLayoutPanelheader.Controls.Add(lblStock_oh, 1, 0);
        //    tableLayoutPanelheader.Controls.Add(lblStock_intransit, 2, 0);
        //    tableLayoutPanelheader.Controls.Add(lblStock, 3, 0);
        //    tableLayoutPanelheader.Controls.Add(ctrlHeader, 4, 0);
        //    pnlHeader.Controls.Add(tableLayoutPanelheader);
        //    TableLayoutPanel tableLayoutPanel1 = new TableLayoutPanel();
        //    tableLayoutPanel1.Width = panel2.Width;
        //    tableLayoutPanel1.Height = panel2.Height;

        //    tableLayoutPanel1.Font = new Font("Verdana", 8, FontStyle.Regular);
        //    //tableLayoutPanel1.CellBorderStyle = CellBorderStyle.FixedSingle;
        //    tableLayoutPanel1.AutoScroll = true;
        //    tableLayoutPanel1.ColumnCount = 5;
        //    for (int i = 0; i < tableLayoutPanel1.ColumnCount; i++)
        //    {
        //        if (i == 0)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 130);
        //            tableLayoutPanel1.ColumnStyles.Add(cs);
        //        }
        //        //if (i == 1)
        //        //{
        //        //    ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 50);
        //        //    tableLayoutPanel1.ColumnStyles.Add(cs);
        //        //}
        //        if (i == 1)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 70);
        //            tableLayoutPanel1.ColumnStyles.Add(cs);
        //        }
        //        if (i == 2)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 70);
        //            tableLayoutPanel1.ColumnStyles.Add(cs);
        //        }

        //        if (i == 3)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 70);
        //            tableLayoutPanel1.ColumnStyles.Add(cs);
        //        }


        //        if (i == 4)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 1515);
        //            tableLayoutPanel1.ColumnStyles.Add(cs);
        //        }

        //    }
        //    if (dsPart != null)
        //    {
        //        //dataGridView1.DataSource = dsPart.Tables[0];
        //        for (int j = 0; j < panel2.Controls.Count; j++)
        //        {
        //            panel2.Controls.RemoveAt(j);
        //        }
                
        //        if (dsCommitReq != null && dsCommitReq.Tables.Count > 0 && dsCommitReq.Tables[0].Rows.Count > 0)
        //        {
        //            lblCommitId.Text = dsCommitReq.Tables[0].Rows[0]["com_id"].ToString();
        //        }

        //        DataView view = new DataView(dsPart.Tables[0]);
        //        distinctValues = view.ToTable(true, "part_number", "short_desc", "Stock_oh", "Stock_transit", "Stock", "lead_time_air", "lead_time_ocean", "lead_time_ground", "std_ship_meth", "plan_type", "ord_multiple", "curr_po_at", "curr_line_at", "curr_po_nl", "curr_line_nl");
        //        int partcount = distinctValues.Rows.Count;
        //        String[] leadtime;
        //        for (int rownum = 0; rownum < partcount; rownum++)  //partcount
        //        {
        //            // DataRow[] drDemand = dsPart.Tables[0].Select("part_num='" + distinctValues.Rows[rownum]["part_num"].ToString() + "'");
        //            //DataRow[] drSupply = dsPart.Tables[1].Select("part_num='" + distinctValues.Rows[rownum]["part_num"].ToString() + "'");
        //            DataRow[] drDemand = dsPart.Tables[0].Select("part_number='" + distinctValues.Rows[rownum]["part_number"].ToString() + "' and SUM_LEVEL='D'");
        //            DataRow[] drSupply = dsPart.Tables[0].Select("part_number='" + distinctValues.Rows[rownum]["part_number"].ToString() + "' and SUM_LEVEL='S'");
        //            DataRow[] drPos = dsPart.Tables[0].Select("part_number='" + distinctValues.Rows[rownum]["part_number"].ToString() + "' and SUM_LEVEL='V'");
        //            DataRow[] drCommit = dsCommitReq.Tables[0].Select("part_num='" + distinctValues.Rows[rownum]["part_number"].ToString() + "'");
        //            DataRow[] drOffset = dsOffset.Tables[0].Select("part_num='" + distinctValues.Rows[rownum]["part_number"].ToString() + "'");

        //            customcontrol.UserControl1 ctrl = new UserControl1();

        //            ctrl.dtOffsetWeek = drOffset.CopyToDataTable();
        //            int[] wk_num;
        //            wk_num = new int[21];
        //            DateTime[] date_wk = new DateTime[21];
        //            string[] date_wk_string = new string[21];
        //            string[] prevdate_wk_string = new string[5];

        //            for (int i = 5; i < dsWeek.Tables[0].Rows.Count; i++)
        //            {
        //                //   MessageBox.Show(dsWeek.Tables[0].Rows[i]["mweek"].ToString());

        //                wk_num[i - 5] = Convert.ToInt32(dsWeek.Tables[0].Rows[i]["mweek"].ToString());
        //                date_wk[i - 5] = Convert.ToDateTime(dsWeek.Tables[0].Rows[i]["date"].ToString());
        //                date_wk_string[i - 5] = dsWeek.Tables[0].Rows[i]["date"].ToString();
        //            }


        //            for (int i = 0; i < 5; i++)
        //            {
        //                //   MessageBox.Show(dsWeek.Tables[0].Rows[i]["mweek"].ToString());

        //                // wk_num[i - 5] = Convert.ToInt32(dsWeek.Tables[0].Rows[i]["mweek"].ToString());
        //                //date_wk[i - 5] = Convert.ToDateTime(dsWeek.Tables[0].Rows[i]["date"].ToString());
        //                prevdate_wk_string[i] = dsWeek.Tables[0].Rows[i]["date"].ToString();
        //            }
                    
        //            ctrl.prevdate_wk_string = prevdate_wk_string;
        //            ctrl.wk = wk_num;
        //            ctrl.date_wk = date_wk;
        //            ctrl.date_wk_string = date_wk_string;

        //      /*   }*/

        //                 int prev_supply = 0;
        //                              string prev_ship_method = "";
        //                              if (drSupply.Length > 0)
        //                              {


        //                                  prev_ship_method = drSupply[0]["prev_ship_method_prev_desc"].ToString();

        //                                  //ctrl.lblPrevShipMethod_04.Text = drSupply[0]["prev_ship_method_prev_4_desc"].ToString();
        //                                  //ctrl.lblPrevShipMethod_03.Text = drSupply[0]["prev_ship_method_prev_3_desc"].ToString();
        //                                  //ctrl.lblPrevShipMethod_02.Text = drSupply[0]["prev_ship_method_prev_2_desc"].ToString();
        //                                  //ctrl.lblPrevShipMethod_01.Text = drSupply[0]["prev_ship_method_prev_1_desc"].ToString();

        //                                  if (drSupply[0]["prev_ship_method_prev_4_desc"].ToString() == "AIR")
        //                                  {
        //                                      ctrl.radioAirwk_04.Checked = true;

        //                                  }
        //                                  else if (drSupply[0]["prev_ship_method_prev_4_desc"].ToString() == "OCEAN")
        //                                  {
        //                                      ctrl.radioOceanWk_04.Checked = true;
        //                                  }
        //                                  else if (drSupply[0]["prev_ship_method_prev_4_desc"].ToString() == "GROUND")
        //                                  {
        //                                      ctrl.radioGroundWk_04.Checked = true;
        //                                  }

        //                                  if (drSupply[0]["prev_ship_method_prev_3_desc"].ToString() == "AIR")
        //                                  {
        //                                      ctrl.radioAirwk_03.Checked = true;

        //                                  }
        //                                  else if (drSupply[0]["prev_ship_method_prev_3_desc"].ToString() == "OCEAN")
        //                                  {
        //                                      ctrl.radioOceanWk_03.Checked = true;
        //                                  }
        //                                  else if (drSupply[0]["prev_ship_method_prev_3_desc"].ToString() == "GROUND")
        //                                  {
        //                                      ctrl.radioGroundWk_03.Checked = true;
        //                                  }

        //                                  if (drSupply[0]["prev_ship_method_prev_2_desc"].ToString() == "AIR")
        //                                  {
        //                                      ctrl.radioAirwk_02.Checked = true;

        //                                  }
        //                                  else if (drSupply[0]["prev_ship_method_prev_2_desc"].ToString() == "OCEAN")
        //                                  {
        //                                      ctrl.radioOceanWk_02.Checked = true;
        //                                  }
        //                                  else if (drSupply[0]["prev_ship_method_prev_2_desc"].ToString() == "GROUND")
        //                                  {
        //                                      ctrl.radioGroundWk_02.Checked = true;
        //                                  }

        //                                  if (drSupply[0]["prev_ship_method_prev_1_desc"].ToString() == "AIR")
        //                                  {
        //                                      ctrl.radioAirwk_01.Checked = true;

        //                                  }
        //                                  else if (drSupply[0]["prev_ship_method_prev_1_desc"].ToString() == "OCEAN")
        //                                  {
        //                                      ctrl.radioOceanWk_01.Checked = true;
        //                                  }
        //                                  else if (drSupply[0]["prev_ship_method_prev_1_desc"].ToString() == "GROUND")
        //                                  {
        //                                      ctrl.radioGroundWk_01.Checked = true;
        //                                  }

        //                                  if (drSupply[0]["prev_ship_method_prev_desc"].ToString() == "AIR")
        //                                  {
        //                                      ctrl.radioAirwk_0.Checked = true;

        //                                  }
        //                                  else if (drSupply[0]["prev_ship_method_prev_desc"].ToString() == "OCEAN")
        //                                  {
        //                                      ctrl.radioOceanWk_0.Checked = true;
        //                                  }
        //                                  else if (drSupply[0]["prev_ship_method_prev_desc"].ToString() == "GROUND")
        //                                  {
        //                                      ctrl.radioGroundWk_0.Checked = true;
        //                                  }


        //                                  ctrl.lblPrevShipMethod_0.Text = drSupply[0]["prev_ship_method_prev_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_1.Text = drSupply[0]["prev_ship_method_curr_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_2.Text = drSupply[0]["prev_ship_method_curr_1_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_3.Text = drSupply[0]["prev_ship_method_curr_2_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_4.Text = drSupply[0]["prev_ship_method_curr_3_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_5.Text = drSupply[0]["prev_ship_method_curr_4_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_6.Text = drSupply[0]["prev_ship_method_curr_5_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_7.Text = drSupply[0]["prev_ship_method_curr_6_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_8.Text = drSupply[0]["prev_ship_method_curr_7_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_9.Text = drSupply[0]["prev_ship_method_curr_8_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_10.Text = drSupply[0]["prev_ship_method_curr_9_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_11.Text = drSupply[0]["prev_ship_method_curr_10_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_12.Text = drSupply[0]["prev_ship_method_curr_11_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_13.Text = drSupply[0]["prev_ship_method_curr_12_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_14.Text = drSupply[0]["prev_ship_method_curr_13_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_15.Text = drSupply[0]["prev_ship_method_curr_14_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_16.Text = drSupply[0]["prev_ship_method_curr_15_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_17.Text = drSupply[0]["prev_ship_method_curr_16_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_18.Text = drSupply[0]["prev_ship_method_curr_17_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_19.Text = drSupply[0]["prev_ship_method_curr_18_desc"].ToString();
        //                                  ctrl.lblPrevShipMethod_20.Text = drSupply[0]["prev_ship_method_curr_19_desc"].ToString();

        //                                  prev_supply = Convert.ToInt32(drSupply[0]["prev_week"].ToString());
        //                                  ctrl.txtSup_wk04.Text = drSupply[0]["prev_week_4"].ToString();
        //                                  ctrl.txtSup_wk03.Text = drSupply[0]["prev_week_3"].ToString();
        //                                  ctrl.txtSup_wk02.Text = drSupply[0]["prev_week_2"].ToString();
        //                                  ctrl.txtSup_wk01.Text = drSupply[0]["prev_week_1"].ToString();


        //                                  ctrl.txtSup_wk0.Text = drSupply[0]["prev_week"].ToString();
        //                                  ctrl.txtSup_wk1.Text = drSupply[0]["cur_week"].ToString();
        //                                  ctrl.txtSup_wk2.Text = drSupply[0]["cur_week_plus_1"].ToString();
        //                                  ctrl.txtSup_wk3.Text = drSupply[0]["cur_week_plus_2"].ToString();
        //                                  ctrl.txtSup_wk4.Text = drSupply[0]["cur_week_plus_3"].ToString();
        //                                  ctrl.txtSup_wk5.Text = drSupply[0]["cur_week_plus_4"].ToString();
        //                                  ctrl.txtSup_wk6.Text = drSupply[0]["cur_week_plus_5"].ToString();
        //                                  ctrl.txtSup_wk7.Text = drSupply[0]["cur_week_plus_6"].ToString();
        //                                  ctrl.txtSup_wk8.Text = drSupply[0]["cur_week_plus_7"].ToString();
        //                                  ctrl.txtSup_wk9.Text = drSupply[0]["cur_week_plus_8"].ToString();
        //                                  ctrl.txtSup_wk10.Text = drSupply[0]["cur_week_plus_9"].ToString();
        //                                  ctrl.txtSup_wk11.Text = drSupply[0]["cur_week_plus_10"].ToString();
        //                                  ctrl.txtSup_wk12.Text = drSupply[0]["cur_week_plus_11"].ToString();
        //                                  ctrl.txtSup_wk13.Text = drSupply[0]["cur_week_plus_12"].ToString();
        //                                  ctrl.txtSup_wk14.Text = drSupply[0]["cur_week_plus_13"].ToString();
        //                                  ctrl.txtSup_wk15.Text = drSupply[0]["cur_week_plus_14"].ToString();
        //                                  ctrl.txtSup_wk16.Text = drSupply[0]["cur_week_plus_15"].ToString();
        //                                  ctrl.txtSup_wk17.Text = drSupply[0]["cur_week_plus_16"].ToString();
        //                                  ctrl.txtSup_wk18.Text = drSupply[0]["cur_week_plus_17"].ToString();
        //                                  ctrl.txtSup_wk19.Text = drSupply[0]["cur_week_plus_18"].ToString();
        //                                  ctrl.txtSup_wk20.Text = drSupply[0]["cur_week_plus_19"].ToString();
        //                              }

        //                              if (drPos.Length > 0)
        //                              {
        //                                  if (ctrl.lblPrevShipMethod_04.Text == "")
        //                                      ctrl.lblPrevShipMethod_04.Text = drPos[0]["prev_ship_method_prev_4_desc"].ToString();
        //                                  if (ctrl.lblPrevShipMethod_03.Text == "")
        //                                      ctrl.lblPrevShipMethod_03.Text = drPos[0]["prev_ship_method_prev_3_desc"].ToString();
        //                                  if (ctrl.lblPrevShipMethod_02.Text == "")
        //                                      ctrl.lblPrevShipMethod_02.Text = drPos[0]["prev_ship_method_prev_2_desc"].ToString();
        //                                  if (ctrl.lblPrevShipMethod_01.Text == "")
        //                                      ctrl.lblPrevShipMethod_01.Text = drPos[0]["prev_ship_method_prev_1_desc"].ToString();
        //                                  if (ctrl.lblPrevShipMethod_0.Text == "")
        //                                      ctrl.lblPrevShipMethod_0.Text = drPos[0]["prev_ship_method_prev_desc"].ToString();

        //                                  ctrl.txtPos_wk1.Text = drPos[0]["cur_week"].ToString();
        //                                  ctrl.txtPos_wk2.Text = drPos[0]["cur_week_plus_1"].ToString();
        //                                  ctrl.txtPos_wk3.Text = drPos[0]["cur_week_plus_2"].ToString();
        //                                  ctrl.txtPos_wk4.Text = drPos[0]["cur_week_plus_3"].ToString();
        //                                  ctrl.txtPos_wk5.Text = drPos[0]["cur_week_plus_4"].ToString();
        //                                  ctrl.txtPos_wk6.Text = drPos[0]["cur_week_plus_5"].ToString();
        //                                  ctrl.txtPos_wk7.Text = drPos[0]["cur_week_plus_6"].ToString();
        //                                  ctrl.txtPos_wk8.Text = drPos[0]["cur_week_plus_7"].ToString();
        //                                  ctrl.txtPos_wk9.Text = drPos[0]["cur_week_plus_8"].ToString();
        //                                  ctrl.txtPos_wk10.Text = drPos[0]["cur_week_plus_9"].ToString();
        //                                  ctrl.txtPos_wk11.Text = drPos[0]["cur_week_plus_10"].ToString();
        //                                  ctrl.txtPos_wk12.Text = drPos[0]["cur_week_plus_11"].ToString();
        //                                  ctrl.txtPos_wk13.Text = drPos[0]["cur_week_plus_12"].ToString();
        //                                  ctrl.txtPos_wk14.Text = drPos[0]["cur_week_plus_13"].ToString();
        //                                  ctrl.txtPos_wk15.Text = drPos[0]["cur_week_plus_14"].ToString();
        //                                  ctrl.txtPos_wk16.Text = drPos[0]["cur_week_plus_15"].ToString();
        //                                  ctrl.txtPos_wk17.Text = drPos[0]["cur_week_plus_16"].ToString();
        //                                  ctrl.txtPos_wk18.Text = drPos[0]["cur_week_plus_17"].ToString();
        //                                  ctrl.txtPos_wk19.Text = drPos[0]["cur_week_plus_18"].ToString();
        //                                  ctrl.txtPos_wk20.Text = drPos[0]["cur_week_plus_19"].ToString();
        //                                  // SET Prev ship method

        //                              }

        //                              if (drCommit.Length > 0 && drCommit.Length >= 24)
        //                              {
        //                                  ctrl.txtSup_wk04.Text = drCommit[0]["prev_com_qty"].ToString();
        //                                  ctrl.txtSup_wk03.Text = drCommit[1]["prev_com_qty"].ToString();
        //                                  ctrl.txtSup_wk02.Text = drCommit[2]["prev_com_qty"].ToString();
        //                                  ctrl.txtSup_wk01.Text = drCommit[3]["prev_com_qty"].ToString();
        //                                  ctrl.txtSup_wk0.Text = drCommit[4]["prev_com_qty"].ToString();

        //                                  ctrl.txtNewSupp_wk1.Text = drCommit[5]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk2.Text = drCommit[6]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk3.Text = drCommit[7]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk4.Text = drCommit[8]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk5.Text = drCommit[9]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk6.Text = drCommit[10]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk7.Text = drCommit[11]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk8.Text = drCommit[12]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk9.Text = drCommit[13]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk10.Text = drCommit[14]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk11.Text = drCommit[15]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk12.Text = drCommit[16]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk13.Text = drCommit[17]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk14.Text = drCommit[18]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk15.Text = drCommit[19]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk16.Text = drCommit[20]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk17.Text = drCommit[21]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk18.Text = drCommit[22]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk19.Text = drCommit[23]["new_req_qty"].ToString();
        //                                  ctrl.txtNewSupp_wk20.Text = "0";
                        
        //                                  if (drCommit[0]["prev_ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_04.Checked = true;

        //                                  }
        //                                  else if (drCommit[0]["prev_ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_04.Checked = true;
        //                                  }
        //                                  else if (drCommit[0]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_04.Checked = true;
        //                                  }

        //                                  if (drCommit[1]["prev_ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_03.Checked = true;

        //                                  }
        //                                  else if (drCommit[1]["prev_ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_03.Checked = true;
        //                                  }
        //                                  else if (drCommit[1]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_03.Checked = true;
        //                                  }


        //                                  if (drCommit[2]["prev_ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_02.Checked = true;

        //                                  }
        //                                  else if (drCommit[2]["prev_ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_02.Checked = true;
        //                                  }
        //                                  else if (drCommit[2]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_02.Checked = true;
        //                                  }

        //                                  if (drCommit[3]["prev_ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_01.Checked = true;

        //                                  }
        //                                  else if (drCommit[3]["prev_ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_01.Checked = true;
        //                                  }
        //                                  else if (drCommit[3]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_03.Checked = true;
        //                                  }

        //                                  if (drCommit[4]["prev_ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_0.Checked = true;

        //                                  }
        //                                  else if (drCommit[4]["prev_ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_0.Checked = true;
        //                                  }
        //                                  else if (drCommit[4]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_0.Checked = true;
        //                                  }


        //                                  //****************************new week********************************
        //                                  if (drCommit[5]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_1.Checked = true;

        //                                  }
        //                                  else if (drCommit[5]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_1.Checked = true;
        //                                  }
        //                                  else if (drCommit[5]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_1.Checked = true;
        //                                  }

        //                                  //wwek#2  
        //                                  if (drCommit[6]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_2.Checked = true;

        //                                  }
        //                                  else if (drCommit[6]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanwk_2.Checked = true;
        //                                  }
        //                                  else if (drCommit[6]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_2.Checked = true;
        //                                  }

        //                                  //wwek#3

        //                                  if (drCommit[7]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_3.Checked = true;

        //                                  }
        //                                  else if (drCommit[7]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_3.Checked = true;
        //                                  }
        //                                  else if (drCommit[7]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_3.Checked = true;
        //                                  }
        //                                  //wwek#2

        //                                  if (drCommit[8]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_4.Checked = true;

        //                                  }
        //                                  else if (drCommit[8]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_4.Checked = true;
        //                                  }
        //                                  else if (drCommit[8]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_4.Checked = true;
        //                                  }
        //                                  //wwek#2

        //                                  if (drCommit[9]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_5.Checked = true;

        //                                  }
        //                                  else if (drCommit[9]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_5.Checked = true;

        //                                  }
        //                                  else if (drCommit[9]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_5.Checked = true;
        //                                  }

        //                                  //wwek#2

        //                                  if (drCommit[10]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_6.Checked = true;

        //                                  }
        //                                  else if (drCommit[10]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_6.Checked = true;
        //                                  }
        //                                  else if (drCommit[10]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_6.Checked = true;
        //                                  }

        //                                  //wwek#2
        //                                  if (drCommit[11]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_7.Checked = true;

        //                                  }
        //                                  else if (drCommit[11]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_7.Checked = true;
        //                                  }
        //                                  else if (drCommit[11]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_7.Checked = true;
        //                                  }

        //                                  //wwek#2

        //                                  if (drCommit[12]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_8.Checked = true;

        //                                  }
        //                                  else if (drCommit[12]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_8.Checked = true;
        //                                  }
        //                                  else if (drCommit[12]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_8.Checked = true;
        //                                  }

        //                                  //wwek#2
        //                                  if (drCommit[13]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_9.Checked = true;

        //                                  }
        //                                  else if (drCommit[13]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_9.Checked = true;
        //                                  }
        //                                  else if (drCommit[13]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_9.Checked = true;
        //                                  }
        //                                  //wwek#2
        //                                  if (drCommit[14]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_10.Checked = true;

        //                                  }
        //                                  else if (drCommit[14]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_10.Checked = true;

        //                                  }
        //                                  else if (drCommit[14]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_10.Checked = true;
        //                                  }

        //                                  //wwek#2
        //                                  if (drCommit[15]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_11.Checked = true;

        //                                  }
        //                                  else if (drCommit[15]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_11.Checked = true;

        //                                  }
        //                                  else if (drCommit[15]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_11.Checked = true;
        //                                  }

        //                                  //wwek#2
        //                                  if (drCommit[16]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_12.Checked = true;

        //                                  }
        //                                  else if (drCommit[16]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_12.Checked = true;
        //                                  }
        //                                  else if (drCommit[16]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_12.Checked = true;
        //                                  }

        //                                  //wwek#2
        //                                  if (drCommit[17]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_13.Checked = true;

        //                                  }
        //                                  else if (drCommit[17]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_13.Checked = true;
        //                                  }
        //                                  else if (drCommit[17]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_13.Checked = true;
        //                                  }


        //                                  //wwek#2
        //                                  if (drCommit[18]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_14.Checked = true;

        //                                  }
        //                                  else if (drCommit[18]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_14.Checked = true;
        //                                  }
        //                                  else if (drCommit[18]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_14.Checked = true;
        //                                  }

        //                                  //wwek#2
        //                                  if (drCommit[19]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_15.Checked = true;

        //                                  }
        //                                  else if (drCommit[19]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_15.Checked = true;

        //                                  }
        //                                  else if (drCommit[19]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_15.Checked = true;
        //                                  }

        //                                  //wwek#2
        //                                  if (drCommit[20]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_16.Checked = true;

        //                                  }
        //                                  else if (drCommit[20]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_16.Checked = true;
        //                                  }
        //                                  else if (drCommit[20]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_16.Checked = true;
        //                                  }

        //                                  //wwek#2
        //                                  if (drCommit[21]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_17.Checked = true;

        //                                  }
        //                                  else if (drCommit[21]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_17.Checked = true;
        //                                  }
        //                                  else if (drCommit[21]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_17.Checked = true;
        //                                  }

        //                                  //wwek#2

        //                                  if (drCommit[22]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_18.Checked = true;

        //                                  }
        //                                  else if (drCommit[22]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_18.Checked = true;
        //                                  }
        //                                  else if (drCommit[22]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_18.Checked = true;
        //                                  }


        //                                  //wwek#2
        //                                  if (drCommit[23]["ship_method"].ToString() == "A")
        //                                  {
        //                                      ctrl.radioAirwk_19.Checked = true;

        //                                  }
        //                                  else if (drCommit[23]["ship_method"].ToString() == "O")
        //                                  {
        //                                      ctrl.radioOceanWk_19.Checked = true;
        //                                  }
        //                                  else if (drCommit[23]["prev_ship_method"].ToString() == "G")
        //                                  {
        //                                      ctrl.radioGroundWk_19.Checked = true;
        //                                  }

        //                                  //wwek#2
        //                                  //if (drCommit[24]["ship_method"].ToString() == "A")
        //                                  //{
        //                                  //    ctrl.radioAirwk_20.Checked = true;

        //                                  //}
        //                                  //else if (drCommit[24]["ship_method"].ToString() == "O")
        //                                  //{
        //                                  //    ctrl.radioOceanWk_20.Checked = true;

        //                                  //}
        //                                  ctrl.radioAirwk_2.CheckedChanged += new EventHandler(ctrl.radioAirwk_2_CheckedChanged);


        //                              }
        //                              else
        //                              {
        //                                  if (distinctValues.Rows[rownum]["std_ship_meth"].ToString() == "AIR")
        //                                  {
        //                                      ctrl.radioAirwk_1.Checked = true;
        //                                      ctrl.radioAirwk_2.Checked = true;
        //                                      ctrl.radioAirwk_3.Checked = true;
        //                                      ctrl.radioAirwk_4.Checked = true;
        //                                      ctrl.radioAirwk_5.Checked = true;
        //                                      ctrl.radioAirwk_6.Checked = true;
        //                                      ctrl.radioAirwk_7.Checked = true;
        //                                      ctrl.radioAirwk_8.Checked = true;
        //                                      ctrl.radioAirwk_9.Checked = true;
        //                                      ctrl.radioAirwk_10.Checked = true;
        //                                      ctrl.radioAirwk_11.Checked = true;
        //                                      ctrl.radioAirwk_12.Checked = true;
        //                                      ctrl.radioAirwk_13.Checked = true;
        //                                      ctrl.radioAirwk_14.Checked = true;
        //                                      ctrl.radioAirwk_15.Checked = true;
        //                                      ctrl.radioAirwk_16.Checked = true;
        //                                      ctrl.radioAirwk_17.Checked = true;
        //                                      ctrl.radioAirwk_18.Checked = true;
        //                                      ctrl.radioAirwk_19.Checked = true;
        //                                      ctrl.radioAirwk_20.Checked = true;

        //                                  }
        //                                  else if (distinctValues.Rows[rownum]["std_ship_meth"].ToString() == "OCEAN")
        //                                  {
        //                                      ctrl.radioOceanWk_1.Checked = true;
        //                                      ctrl.radioOceanwk_2.Checked = true;
        //                                      ctrl.radioOceanWk_3.Checked = true;
        //                                      ctrl.radioOceanWk_4.Checked = true;
        //                                      ctrl.radioOceanWk_5.Checked = true;
        //                                      ctrl.radioOceanWk_6.Checked = true;
        //                                      ctrl.radioOceanWk_7.Checked = true;
        //                                      ctrl.radioOceanWk_8.Checked = true;
        //                                      ctrl.radioOceanWk_9.Checked = true;
        //                                      ctrl.radioOceanWk_10.Checked = true;
        //                                      ctrl.radioOceanWk_11.Checked = true;
        //                                      ctrl.radioOceanWk_12.Checked = true;
        //                                      ctrl.radioOceanWk_13.Checked = true;
        //                                      ctrl.radioOceanWk_14.Checked = true;
        //                                      ctrl.radioOceanWk_15.Checked = true;
        //                                      ctrl.radioOceanWk_16.Checked = true;
        //                                      ctrl.radioOceanWk_17.Checked = true;
        //                                      ctrl.radioOceanWk_18.Checked = true;
        //                                      ctrl.radioOceanWk_19.Checked = true;
        //                                      ctrl.radioOceanWk_20.Checked = true;

        //                                  }
        //                                  else if (distinctValues.Rows[rownum]["std_ship_meth"].ToString() == "GROUND")
        //                                  {
        //                                      ctrl.radioGroundWk_1.Checked = true;
        //                                      ctrl.radioGroundWk_2.Checked = true;
        //                                      ctrl.radioGroundWk_3.Checked = true;
        //                                      ctrl.radioGroundWk_4.Checked = true;
        //                                      ctrl.radioGroundWk_5.Checked = true;
        //                                      ctrl.radioGroundWk_6.Checked = true;
        //                                      ctrl.radioGroundWk_7.Checked = true;
        //                                      ctrl.radioGroundWk_8.Checked = true;
        //                                      ctrl.radioGroundWk_9.Checked = true;
        //                                      ctrl.radioGroundWk_10.Checked = true;
        //                                      ctrl.radioGroundWk_11.Checked = true;
        //                                      ctrl.radioGroundWk_12.Checked = true;
        //                                      ctrl.radioGroundWk_13.Checked = true;
        //                                      ctrl.radioGroundWk_14.Checked = true;
        //                                      ctrl.radioGroundWk_15.Checked = true;
        //                                      ctrl.radioGroundWk_16.Checked = true;
        //                                      ctrl.radioGroundWk_17.Checked = true;
        //                                      ctrl.radioGroundWk_18.Checked = true;
        //                                      ctrl.radioGroundWk_19.Checked = true;
        //                                      ctrl.radioGroundWk_20.Checked = true;

        //                                  }

        //                              }
        //                              if (drDemand.Length > 0)
        //                              {
        //                                  ctrl.textBox6.Text = drDemand[0]["cur_week"].ToString();
        //                                  ctrl.txtDem_wk_1.Text = drDemand[0]["cur_week"].ToString();
        //                                  ctrl.txtDem_wk_2.Text = drDemand[0]["cur_week_plus_1"].ToString();
        //                                  ctrl.txtDem_wk_3.Text = drDemand[0]["cur_week_plus_2"].ToString();
        //                                  ctrl.txtDem_wk_4.Text = drDemand[0]["cur_week_plus_3"].ToString();
        //                                  ctrl.txtDem_wk_5.Text = drDemand[0]["cur_week_plus_4"].ToString();
        //                                  ctrl.txtDem_wk_6.Text = drDemand[0]["cur_week_plus_5"].ToString();
        //                                  ctrl.txtDem_wk_7.Text = drDemand[0]["cur_week_plus_6"].ToString();
        //                                  ctrl.txtDem_wk_8.Text = drDemand[0]["cur_week_plus_7"].ToString();
        //                                  ctrl.txtDem_wk_9.Text = drDemand[0]["cur_week_plus_8"].ToString();
        //                                  ctrl.txtDem_wk_10.Text = drDemand[0]["cur_week_plus_9"].ToString();
        //                                  ctrl.txtDem_wk_11.Text = drDemand[0]["cur_week_plus_10"].ToString();
        //                                  ctrl.txtDem_wk_12.Text = drDemand[0]["cur_week_plus_11"].ToString();
        //                                  ctrl.txtDem_wk_13.Text = drDemand[0]["cur_week_plus_12"].ToString();
        //                                  ctrl.txtDem_wk_14.Text = drDemand[0]["cur_week_plus_13"].ToString();
        //                                  ctrl.txtDem_wk_15.Text = drDemand[0]["cur_week_plus_14"].ToString();
        //                                  ctrl.txtDem_wk_16.Text = drDemand[0]["cur_week_plus_15"].ToString();
        //                                  ctrl.txtDem_wk_17.Text = drDemand[0]["cur_week_plus_16"].ToString();
        //                                  ctrl.txtDem_wk_18.Text = drDemand[0]["cur_week_plus_17"].ToString();
        //                                  ctrl.txtDem_wk_19.Text = drDemand[0]["cur_week_plus_18"].ToString();
        //                                  ctrl.txtDem_wk_20.Text = drDemand[0]["cur_week_plus_19"].ToString();

        //                              }
        //            ctrl.part_stock = Convert.ToInt32(distinctValues.Rows[rownum]["Stock_oh"].ToString());
        //            ctrl.part_stock = Convert.ToInt32(distinctValues.Rows[rownum]["Stock_transit"].ToString());
        //            ctrl.part_stock = Convert.ToInt32(distinctValues.Rows[rownum]["Stock"].ToString());
        //            ctrl.air_lead_time = Convert.ToInt32(distinctValues.Rows[rownum]["lead_time_air"].ToString());
        //          //  if (distinctValues.Rows[rownum]["lead_time_ground"].ToString()!="")
        //            ctrl.ground_lead_time = Convert.ToInt32(distinctValues.Rows[rownum]["lead_time_ground"].ToString());
        //            ctrl.ocean_lead_time = Convert.ToInt32(distinctValues.Rows[rownum]["lead_time_ocean"].ToString());
        //            ctrl.lblPlanType.Text = distinctValues.Rows[rownum]["plan_type"].ToString();
        //            ctrl.lblOrdMult.Text = distinctValues.Rows[rownum]["ord_multiple"].ToString();
        //            ctrl.lblStdShipMeth.Text = distinctValues.Rows[rownum]["std_ship_meth"].ToString();


        //            if (cmbSite.SelectedValue.ToString() == "AT")
        //            {
        //                ctrl.lblPONum.Text = distinctValues.Rows[rownum]["curr_po_at"].ToString();
        //                ctrl.lblLineNum.Text = distinctValues.Rows[rownum]["curr_line_at"].ToString();
        //            }
        //            else
        //            {
        //                ctrl.lblPONum.Text = distinctValues.Rows[rownum]["curr_po_nl"].ToString();
        //                ctrl.lblLineNum.Text = distinctValues.Rows[rownum]["curr_line_nl"].ToString();
        //            }

        //            int weekOffset;
        //          if (prev_ship_method == "AIR")
        //                weekOffset = getoffset(date_wk[0], Convert.ToInt32(distinctValues.Rows[rownum]["lead_time_air"].ToString()));
        //            else if (prev_ship_method == "GROUND")
        //                weekOffset = getoffset(date_wk[0], Convert.ToInt32(distinctValues.Rows[rownum]["lead_time_ground"].ToString()));
        //            else  
        //                weekOffset = getoffset(date_wk[0], Convert.ToInt32(distinctValues.Rows[rownum]["lead_time_ocean"].ToString()));

        //            //wk_num[i-1] = Convert.ToInt32(dsWeek.Tables[0].Rows[i]["mweek"].ToString());
        //            //     date_wk[i-1]

        //            int offsetWeekprev = -1;
        //            for (int intwk = 0; intwk < 20; intwk++)
        //            {
        //                if (wk_num[intwk] == weekOffset)
        //                {
        //                    offsetWeekprev = intwk;
        //                    break;
        //                }
        //            }

        //            ctrl.Name = "dem" + distinctValues.Rows[rownum]["part_number"].ToString();
        //            ctrl.Width = 1515;
        //            Label lblPartName = new Label();
        //            lblPartName.Text = distinctValues.Rows[rownum]["part_number"].ToString();
        //            lblPartName.Width = 130;
        //            lblPartName.TextAlign = ContentAlignment.MiddleLeft;
        //            lblPartName.BackColor = System.Drawing.Color.Red;
        //            //Label lblPartDesc = new Label();
        //            //lblPartDesc.Text = distinctValues.Rows[rownum]["short_desc"].ToString();
        //            //lblPartDesc.Width = 50;
        //            Label lblstock_oh = new Label();
        //            lblstock_oh.TextAlign = ContentAlignment.MiddleCenter;
        //            lblstock_oh.Width = 70;
        //            lblstock_oh.BackColor = System.Drawing.Color.Red;
        //            lblstock_oh.Text = distinctValues.Rows[rownum]["Stock_oh"].ToString();
                    
        //            Label lblstock_inTransit = new Label();
        //            lblstock_inTransit.TextAlign = ContentAlignment.MiddleCenter;
        //            lblstock_inTransit.Width = 70;
        //            lblstock_inTransit.BackColor = System.Drawing.Color.Red;
        //            lblstock_inTransit.Text = distinctValues.Rows[rownum]["Stock_transit"].ToString();

        //            Label lblstock = new Label();
        //            lblstock.TextAlign = ContentAlignment.MiddleCenter;
        //            lblstock.Width = 70;
        //            lblstock.BackColor = System.Drawing.Color.Red;
        //            lblstock.Text = distinctValues.Rows[rownum]["Stock"].ToString();


        //            if (rownum % 2 == 0)
        //            {
        //                lblPartName.BackColor = System.Drawing.Color.Bisque;
        //              //  lblPartDesc.BackColor = System.Drawing.Color.Bisque;
        //                lblstock.BackColor = System.Drawing.Color.Bisque;
        //                lblstock_oh.BackColor = System.Drawing.Color.Bisque;
        //                lblstock_inTransit.BackColor = System.Drawing.Color.Bisque;
        //                ctrl.BackColor = System.Drawing.Color.Bisque;
        //            }
        //            else
        //            {
        //                lblPartName.BackColor = System.Drawing.Color.PeachPuff;
        //               // lblPartDesc.BackColor = System.Drawing.Color.PeachPuff;
        //                lblstock.BackColor = System.Drawing.Color.PeachPuff;
        //                lblstock_oh.BackColor = System.Drawing.Color.PeachPuff;
        //                lblstock_inTransit.BackColor = System.Drawing.Color.PeachPuff;
        //                ctrl.BackColor = System.Drawing.Color.PeachPuff;
        //            }
        //            ctrl.getOffsetForAll();
        //            //if (prev_supply > 0)
        //            //{

        //            //    TextBox txt = ctrl.Controls.Find("txtOffset_wk" + offsetWeekprev, true)[0] as TextBox;
        //            //    if (txt.Text != "")
        //            //        txt.Text = (Convert.ToInt32(txt.Text) + prev_supply).ToString();
        //            //    else
        //            //        txt.Text = prev_supply.ToString();
        //            //}

        //            tableLayoutPanel1.Controls.Add(lblPartName, 0, rownum + 1);
        //            // tableLayoutPanel1.Controls.Add(lblPartDesc, 1, rownum + 1);
        //            tableLayoutPanel1.Controls.Add(lblstock_oh, 1, rownum + 1);
        //            tableLayoutPanel1.Controls.Add(lblstock_inTransit, 2, rownum + 1);
        //            tableLayoutPanel1.Controls.Add(lblstock, 3, rownum + 1);
        //            tableLayoutPanel1.Controls.Add(ctrl, 4, rownum + 1);
        //        }
        //        tableLayoutPanel1.CellPaint += new TableLayoutCellPaintEventHandler(tableLayoutPanel1_CellPaint);
        //        panel2.Controls.Add(tableLayoutPanel1);
        //    }

        //    //}
        //    // catch (Exception ex)
        //    //{
        //    //    MessageBox.Show("Error" + ex.Message.ToString());
        //    // }
        //}
        private void tableLayoutPanel1_CellPaint(object sender, TableLayoutCellPaintEventArgs e)
        {
            if (e.Row % 2 == 0)
            {
                e.Graphics.FillRectangle(Brushes.PeachPuff, e.CellBounds);
            }
            else
            {
                e.Graphics.FillRectangle(Brushes.Bisque, e.CellBounds);
            }
        }


        private void cmbSite_SelectedIndexChanged(object sender, EventArgs e)
        {


        }

        //private void btnDownloadGrid_Click(object sender, EventArgs e)
        //{
        //    Cursor.Current = Cursors.WaitCursor;
        //    if (cmbSite.SelectedIndex == 0 || cmbVendor.SelectedIndex == 0)
        //    {
        //        MessageBox.Show("Please select site and vendor ");
        //    }
        //    else
        //    {
        //        getPart();
        //    }
        //    Cursor.Current = Cursors.Default;
        //}

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


        public int getoffset(DateTime date, int lead_time)
        {
            string sql = "exec getOffsetWeek '" + date + "'," + lead_time.ToString();
            try
            {
                SqlConnection cn = new SqlConnection(this.constr);
                cn.Open();
                SqlCommand cmd = new SqlCommand(sql, cn);                
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet dsOffsetWeek = new DataSet();
                da.Fill(dsOffsetWeek);
                return Convert.ToInt32(dsOffsetWeek.Tables[0].Rows[0][0].ToString());

            }
            catch (Exception ex)
            {
                return 0;
            }
        }
        //private void button1_Click(object sender, EventArgs e)
        //{
        //    string sql = "";
        //    int com_id = 0;

        //    for (int r = 0; r < distinctValues.Rows.Count; r++)
        //    {
        //        string part = distinctValues.Rows[r]["part_number"].ToString();

        //        // drDemand = dsPart.Tables[0].Select("part_num='" + distinctValues.Rows[r]["part_num"].ToString() + "'");
        //        // DataRow[] drSupply = dsPart.Tables[1].Select("part_num='" + distinctValues.Rows[r]["part_num"].ToString() + "'");
        //        UserControl1 user = panel2.Controls.Find("dem" + part, true)[0] as UserControl1;
        //        UserControl2 header = this.pnlHeader.Controls.Find("header", true)[0] as UserControl2;
        //        if (user != null)
        //        {

        //            string plan_type = "M";
        //            string po = "0";
        //            string line = "0";

        //            int new_com_qty_1 = 0;
        //            int new_com_qty_2 = 0;
        //            int new_com_qty_3 = 0;
        //            int new_com_qty_4 = 0;
        //            int new_com_qty_5 = 0;
        //            int new_com_qty_6 = 0;
        //            int new_com_qty_7 = 0;
        //            int new_com_qty_8 = 0;
        //            int new_com_qty_9 = 0;
        //            int new_com_qty_10 = 0;
        //            int new_com_qty_11 = 0;
        //            int new_com_qty_12 = 0;
        //            int new_com_qty_13 = 0;
        //            int new_com_qty_14 = 0;
        //            int new_com_qty_15 = 0;
        //            int new_com_qty_16 = 0;
        //            int new_com_qty_17 = 0;
        //            int new_com_qty_18 = 0;
        //            int new_com_qty_19 = 0;
        //            int new_com_qty_20 = 0;
        //            string ship_method1 = "A";
        //            string ship_method2 = "A";
        //            string ship_method3 = "A";
        //            string ship_method4 = "A";
        //            string ship_method5 = "A";
        //            string ship_method6 = "A";
        //            string ship_method7 = "A";
        //            string ship_method8 = "A";
        //            string ship_method9 = "A";
        //            string ship_method10 = "A";
        //            string ship_method11 = "A";
        //            string ship_method12 = "A";
        //            string ship_method13 = "A";
        //            string ship_method14 = "A";
        //            string ship_method15 = "A";
        //            string ship_method16 = "A";
        //            string ship_method17 = "A";
        //            string ship_method18 = "A";
        //            string ship_method19 = "A";
        //            string ship_method20 = "A";

        //            //**********Wk #1 New Request,ship method ***********************
        //            int new_com_week_1 = user.p_wk_1;
        //            if (user.txtNewSupp_wk1.Text != "")
        //            {
        //                new_com_qty_1 = Convert.ToInt32(user.txtNewSupp_wk1.Text);
        //            }
        //            if (user.shipMethod[0].ToString() != "")
        //                ship_method1 = user.shipMethod[0].ToString();

        //            //**********Wk #2 New Request,ship method ***********************
        //            int new_com_week_2 = user.p_wk_2;
        //            if (user.txtNewSupp_wk2.Text != "")
        //            {
        //                new_com_qty_2 = Convert.ToInt32(user.txtNewSupp_wk2.Text);
        //            }
        //            if (user.shipMethod[1].ToString() != "")
        //                ship_method2 = user.shipMethod[1].ToString();

        //            //**********Wk #3 New Request,ship method ***********************
        //            int new_com_week_3 = user.p_wk_3;
        //            if (user.txtNewSupp_wk3.Text != "")
        //                new_com_qty_3 = Convert.ToInt32(user.txtNewSupp_wk3.Text);
        //            if (user.shipMethod[2].ToString() != "")
        //                ship_method3 = user.shipMethod[2].ToString();

        //            //**********Wk #4 New Request,ship method ***********************
        //            int new_com_week_4 = user.p_wk_4;
        //            if (user.txtNewSupp_wk4.Text != "")
        //                new_com_qty_4 = Convert.ToInt32(user.txtNewSupp_wk4.Text);
        //            if (user.shipMethod[3].ToString() != "")
        //                ship_method4 = user.shipMethod[3].ToString();

        //            //**********Wk #5 New Request,ship method ***********************  
        //            int new_com_week_5 = user.p_wk_5;
        //            if (user.txtNewSupp_wk5.Text != "")
        //                new_com_qty_5 = Convert.ToInt32(user.txtNewSupp_wk5.Text);
        //            if (user.shipMethod[4].ToString() != "")
        //                ship_method5 = user.shipMethod[4].ToString();

        //            //**********Wk #5 New Request,ship method ***********************  
        //            int new_com_week_6 = user.p_wk_6;
        //            if (user.txtNewSupp_wk6.Text != "")
        //                new_com_qty_6 = Convert.ToInt32(user.txtNewSupp_wk6.Text);
        //            if (user.shipMethod[5].ToString() != "")
        //                ship_method6 = user.shipMethod[5].ToString();

        //            int new_com_week_7 = user.p_wk_7;
        //            if (user.txtNewSupp_wk7.Text != "")
        //                new_com_qty_7 = Convert.ToInt32(user.txtNewSupp_wk7.Text);
        //            if (user.shipMethod[6].ToString() != "")
        //                ship_method7 = user.shipMethod[6].ToString();

        //            int new_com_week_8 = user.p_wk_8;
        //            if (user.txtNewSupp_wk8.Text != "")
        //                new_com_qty_8 = Convert.ToInt32(user.txtNewSupp_wk8.Text);
        //            if (user.shipMethod[7].ToString() != "")
        //                ship_method8 = user.shipMethod[7].ToString();

        //            int new_com_week_9 = user.p_wk_9;
        //            if (user.txtNewSupp_wk9.Text != "")
        //                new_com_qty_9 = Convert.ToInt32(user.txtNewSupp_wk9.Text);
        //            if (user.shipMethod[8].ToString() != "")
        //                ship_method9 = user.shipMethod[8].ToString();

        //            int new_com_week_10 = user.p_wk_10;
        //            if (user.txtNewSupp_wk10.Text != "")
        //                new_com_qty_10 = Convert.ToInt32(user.txtNewSupp_wk10.Text);
        //            if (user.shipMethod[10].ToString() != "")
        //                ship_method10 = user.shipMethod[10].ToString();

        //            int new_com_week_11 = user.p_wk_11;
        //            if (user.txtNewSupp_wk11.Text != "")
        //                new_com_qty_11 = Convert.ToInt32(user.txtNewSupp_wk11.Text);
        //            if (user.shipMethod[10].ToString() != "")
        //                ship_method11 = user.shipMethod[10].ToString();

        //            int new_com_week_12 = user.p_wk_12;
        //            if (user.txtNewSupp_wk12.Text != "")
        //                new_com_qty_12 = Convert.ToInt32(user.txtNewSupp_wk12.Text);
        //            if (user.shipMethod[11].ToString() != "")
        //                ship_method12 = user.shipMethod[11].ToString();

        //            int new_com_week_13 = user.p_wk_13;
        //            if (user.txtNewSupp_wk13.Text != "")
        //                new_com_qty_13 = Convert.ToInt32(user.txtNewSupp_wk13.Text);
        //            if (user.shipMethod[12].ToString() != "")
        //                ship_method13 = user.shipMethod[12].ToString();

        //            int new_com_week_14 = user.p_wk_14;
        //            if (user.txtNewSupp_wk14.Text != "")
        //                new_com_qty_14 = Convert.ToInt32(user.txtNewSupp_wk14.Text);
        //            if (user.shipMethod[13].ToString() != "")
        //                ship_method14 = user.shipMethod[13].ToString();

        //            int new_com_week_15 = user.p_wk_15;
        //            if (user.txtNewSupp_wk15.Text != "")
        //                new_com_qty_15 = Convert.ToInt32(user.txtNewSupp_wk15.Text);
        //            if (user.shipMethod[14].ToString() != "")
        //                ship_method15 = user.shipMethod[14].ToString();

        //            int new_com_week_16 = user.p_wk_16;
        //            if (user.txtNewSupp_wk16.Text != "")
        //                new_com_qty_16 = Convert.ToInt32(user.txtNewSupp_wk16.Text);
        //            if (user.shipMethod[15].ToString() != "")
        //                ship_method16 = user.shipMethod[15].ToString();

        //            int new_com_week_17 = user.p_wk_17;
        //            if (user.txtNewSupp_wk5.Text != "")
        //                new_com_qty_17 = Convert.ToInt32(user.txtNewSupp_wk17.Text);
        //            if (user.shipMethod[16].ToString() != "")
        //                ship_method17 = user.shipMethod[16].ToString();

        //            int new_com_week_18 = user.p_wk_18;
        //            if (user.txtNewSupp_wk18.Text != "")
        //                new_com_qty_18 = Convert.ToInt32(user.txtNewSupp_wk18.Text);
        //            if (user.shipMethod[17].ToString() != "")
        //                ship_method18 = user.shipMethod[17].ToString();

        //            int new_com_week_19 = user.p_wk_19;
        //            if (user.txtNewSupp_wk19.Text != "")
        //                new_com_qty_19 = Convert.ToInt32(user.txtNewSupp_wk19.Text);
        //            if (user.shipMethod[18].ToString() != "")
        //                ship_method19 = user.shipMethod[18].ToString();

        //            int new_com_week_20 = user.p_wk_20;
        //            if (user.txtNewSupp_wk20.Text != "")
        //                new_com_qty_20 = Convert.ToInt32(user.txtNewSupp_wk20.Text);
        //            if (user.shipMethod[19].ToString() != "")
        //                ship_method20 = user.shipMethod[19].ToString();
        //            if (user.lblPlanType.Text.ToString() != "")

        //                plan_type = user.lblPlanType.Text.ToString();
        //            po = user.lblPONum.Text.ToString();
        //            if (user.lblLineNum.Text.ToString() != "")
        //                line = user.lblLineNum.Text.ToString();

        //            int prev_week_supply = 0;
        //            int prev_week_supply_1 = 0;
        //            int prev_week_supply_2 = 0;
        //            int prev_week_supply_3 = 0;
        //            int prev_week_supply_4 = 0;

        //            if (user.txtSup_wk0.Text != "")
        //                prev_week_supply = Convert.ToInt32(user.txtSup_wk0.Text);
        //            if (user.txtSup_wk01.Text != "")
        //                prev_week_supply_1 = Convert.ToInt32(user.txtSup_wk01.Text);
        //            if (user.txtSup_wk02.Text != "")
        //                prev_week_supply_2 = Convert.ToInt32(user.txtSup_wk02.Text);
        //            if (user.txtSup_wk03.Text != "")
        //                prev_week_supply_3 = Convert.ToInt32(user.txtSup_wk03.Text);
        //            if (user.txtSup_wk04.Text != "")
        //                prev_week_supply_4 = Convert.ToInt32(user.txtSup_wk04.Text);

        //            string prev_week_ship_method = "A";
        //            string prev_week_ship_method_1 = "A";
        //            string prev_week_ship_method_2 = "A";
        //            string prev_week_ship_method_3 = "A";
        //            string prev_week_ship_method_4 = "A";
        //            if (user.radioOceanWk_04.Checked)
        //                prev_week_ship_method_4 = "O";
        //            if (user.radioOceanWk_03.Checked)
        //                prev_week_ship_method_3 = "O";
        //            if (user.radioOceanWk_02.Checked)
        //                prev_week_ship_method_2 = "O";
        //            if (user.radioOceanWk_01.Checked)
        //                prev_week_ship_method_1 = "O";
        //            if (user.radioOceanWk_0.Checked)
        //                prev_week_ship_method = "O";



        //            //prev_qty_20 = Convert.ToInt32(drSupply[r1][20].ToString());

        //            if (lblCommitId.Text.Replace(" ", "").Trim() != "")
        //            {
        //                com_id = Convert.ToInt32(lblCommitId.Text.Replace(" ", "").Trim());
        //            }
        //            try
        //            {
        //                if (com_id != 0)
        //                {
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id.ToString() + " and part_num='" + part + "' and week=" + this.wk_num_1.ToString() + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue.ToString() + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_1 + "'," + this.wk_num_1.ToString() + "," + this.year_1.ToString() + "," + new_com_qty_1 + ",'" + cmbSite.SelectedValue + "','" + ship_method1 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_1.ToString() + "-" + this.year_1 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";

        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_1 + "',ship_method='" + ship_method1 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate())  where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_1;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_2 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_2 + "'," + this.wk_num_2 + "," + this.year_2 + "," + new_com_qty_2 + ",'" + cmbSite.SelectedValue + "','" + ship_method2 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_2.ToString("00") + "-" + this.year_2 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_2 + "',ship_method='" + ship_method2 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_2;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_3 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_3 + "'," + this.wk_num_3 + "," + this.year_3 + "," + new_com_qty_3 + ",'" + cmbSite.SelectedValue + "','" + ship_method3 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_3.ToString("00") + "-" + this.year_3 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_3 + "',ship_method='" + ship_method3 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_3;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_4 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_4 + "'," + this.wk_num_4 + "," + this.year_4 + "," + new_com_qty_4 + ",'" + cmbSite.SelectedValue + "','" + ship_method4 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_4.ToString("00") + "-" + this.year_4 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_4 + "',ship_method='" + ship_method4 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_4;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_5 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_5 + "'," + this.wk_num_5 + "," + this.year_5 + "," + new_com_qty_5 + ",'" + cmbSite.SelectedValue + "','" + ship_method5 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_5.ToString("00") + "-" + this.year_5 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_5 + "',ship_method='" + ship_method5 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_5;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_6 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_6 + "'," + this.wk_num_6 + "," + this.year_6 + "," + new_com_qty_6 + ",'" + cmbSite.SelectedValue + "','" + ship_method6 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_6.ToString("00") + "-" + this.year_6 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_6 + "',ship_method='" + ship_method6 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_6;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_7 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_7 + "'," + this.wk_num_7 + "," + this.year_7 + "," + new_com_qty_7 + ",'" + cmbSite.SelectedValue + "','" + ship_method7 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_7.ToString("00") + "-" + this.year_7 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_7 + "',ship_method='" + ship_method7 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_7;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_8 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_8 + "'," + this.wk_num_8 + "," + this.year_8 + "," + new_com_qty_8 + ",'" + cmbSite.SelectedValue + "','" + ship_method8 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_8.ToString("00") + "-" + this.year_8 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_8 + "',ship_method='" + ship_method8 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_8;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_9 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_9 + "'," + this.wk_num_9 + "," + this.year_9 + "," + new_com_qty_9 + ",'" + cmbSite.SelectedValue + "','" + ship_method9 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_9.ToString("00") + "-" + this.year_9 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_9 + "',ship_method='" + ship_method9 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where  status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_9;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_10 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_10 + "'," + this.wk_num_10 + "," + this.year_10 + "," + new_com_qty_10 + ",'" + cmbSite.SelectedValue + "','" + ship_method10 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_10.ToString("00") + "-" + this.year_10 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_10 + "',ship_method='" + ship_method10 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_10;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_11 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_11 + "'," + this.wk_num_11 + "," + this.year_11 + "," + new_com_qty_11 + ",'" + cmbSite.SelectedValue + "','" + ship_method11 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_11.ToString("00") + "-" + this.year_11 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_11 + "',ship_method='" + ship_method11 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_11;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_12 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_12 + "'," + this.wk_num_12 + "," + this.year_12 + "," + new_com_qty_12 + ",'" + cmbSite.SelectedValue + "','" + ship_method12 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_12.ToString("00") + "-" + this.year_12 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_12 + "',ship_method='" + ship_method12 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_12;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_13 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_13 + "'," + this.wk_num_13 + "," + this.year_13 + "," + new_com_qty_13 + ",'" + cmbSite.SelectedValue + "','" + ship_method13 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_13.ToString("00") + "-" + this.year_13 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_13 + "',ship_method='" + ship_method13 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_13;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_14 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_14 + "'," + this.wk_num_14 + "," + this.year_14 + "," + new_com_qty_14 + ",'" + cmbSite.SelectedValue + "','" + ship_method14 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_14.ToString("00") + "-" + this.year_14 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_14 + "',ship_method='" + ship_method14 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_14;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_15 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_15 + "'," + this.wk_num_15 + "," + this.year_15 + "," + new_com_qty_15 + ",'" + cmbSite.SelectedValue + "','" + ship_method15 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_15.ToString("00") + "-" + this.year_15 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_15 + "',ship_method='" + ship_method15 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_15;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_16 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_16 + "'," + this.wk_num_16 + "," + this.year_16 + "," + new_com_qty_16 + ",'" + cmbSite.SelectedValue + "','" + ship_method16 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_16.ToString("00") + "-" + this.year_16 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_16 + "',ship_method='" + ship_method16 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_16;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_17 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_17 + "'," + this.wk_num_17 + "," + this.year_17 + "," + new_com_qty_17 + ",'" + cmbSite.SelectedValue + "','" + ship_method17 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_17.ToString("00") + "-" + this.year_17 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_17 + "',ship_method='" + ship_method17 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_17;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_18 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_18 + "'," + this.wk_num_18 + "," + this.year_18 + "," + new_com_qty_18 + ",'" + cmbSite.SelectedValue + "','" + ship_method18 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_18.ToString("00") + "-" + this.year_18 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_18 + "',ship_method='" + ship_method18 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_18;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_19 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_19 + "'," + this.wk_num_19 + "," + this.year_19 + "," + new_com_qty_19 + ",'" + cmbSite.SelectedValue + "','" + ship_method19 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_19.ToString("00") + "-" + this.year_19 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')  else ";
        //                    sql = sql + "  update vc_commit set new_req_qty='" + new_com_qty_19 + "',ship_method='" + ship_method19 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_19;
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_20 + "  if (@cnt=0 ) insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,com_id,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_20 + "'," + this.wk_num_20 + "," + this.year_20 + "," + new_com_qty_20 + ",'" + cmbSite.SelectedValue + "','" + ship_method20 + "','" + com_id + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_20.ToString("00") + "-" + this.year_20 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + " else update vc_commit set new_req_qty='" + new_com_qty_20 + "',ship_method='" + ship_method20 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_20;
        //                    //*****************prev_commit data for week04 ******************************************************************
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_04 + "  if (@cnt=0 ) insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line,com_id) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_04 + "','" + header.v_wk_04 + "','" + header.v_yr_04 + "','" + user.txtSup_wk04.Text + "','" + prev_week_ship_method_4 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_04.ToString("00") + "-" + header.v_yr_04 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "','" + com_id + "') else ";
        //                    sql = sql + "  update vc_commit set prev_com_qty='" + user.txtSup_wk04.Text + "',prev_ship_method='" + prev_week_ship_method_4 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate())  where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_04;
        //                    //*****************prev_commit data for week03 ******************************************************************
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_03 + "  if (@cnt=0 ) insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line,com_id) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_03 + "','" + header.v_wk_03 + "','" + header.v_yr_03 + "','" + user.txtSup_wk03.Text + "','" + prev_week_ship_method_3 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_03.ToString("00") + "-" + header.v_yr_03 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "','" + com_id + "') else ";
        //                    sql = sql + "  update vc_commit set prev_com_qty='" + user.txtSup_wk03.Text + "',prev_ship_method='" + prev_week_ship_method_3 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate())  where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_03;
        //                    //*****************prev_commit data for week02 ******************************************************************
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_02 + "  if (@cnt=0 ) insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line,com_id) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_02 + "','" + header.v_wk_02 + "','" + header.v_yr_02 + "','" + user.txtSup_wk02.Text + "','" + prev_week_ship_method_2 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_02.ToString("00") + "-" + header.v_yr_02 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "','" + com_id + "') else ";
        //                    sql = sql + "  update vc_commit set prev_com_qty='" + user.txtSup_wk02.Text + "',prev_ship_method='" + prev_week_ship_method_2 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate())  where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_02;
        //                    //*****************prev_commit data for week01 ******************************************************************
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_01 + "  if (@cnt=0 ) insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line,com_id) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_01 + "','" + header.v_wk_01 + "','" + header.v_yr_01 + "','" + user.txtSup_wk01.Text + "','" + prev_week_ship_method_1 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_01.ToString("00") + "-" + header.v_yr_01 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "','" + com_id + "') else ";
        //                    sql = sql + "  update vc_commit set prev_com_qty='" + user.txtSup_wk01.Text + "',prev_ship_method='" + prev_week_ship_method_1 + "',create_date=getdate(),create_week=dbo.udf_GetISOWeekNumberFromDate(getdate())  where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_01;
        //                    //*****************prev_commit data for week0 ******************************************************************
        //                    sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_0 + "  if (@cnt=0 ) insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line,com_id) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_0 + "','" + header.v_wk_0 + "','" + header.v_yr_0 + "','" + user.txtSup_wk0.Text + "','" + prev_week_ship_method + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_0.ToString("00") + "-" + header.v_yr_0 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "','"+com_id+"') else ";
        //                    sql = sql + "  update vc_commit set prev_com_qty='" + user.txtSup_wk0.Text + "',prev_ship_method='" + prev_week_ship_method + "',create_date=getdate() ,create_week=dbo.udf_GetISOWeekNumberFromDate(getdate()) where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_0;

        //                    /*  sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_04 + "','" + header.v_wk_04 + "','" + header.v_yr_04 + "','" + user.txtSup_wk04.Text + "','" + user.lblPrevShipMethod_04.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_04 + "-" + header.v_yr_04 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                      sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_03 + "','" + header.v_wk_03 + "','" + header.v_yr_03 + "','" + user.txtSup_wk03.Text + "','" + user.lblPrevShipMethod_03.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_03 + "-" + header.v_yr_03 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                      sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_02 + "','" + header.v_wk_02 + "','" + header.v_yr_02 + "','" + user.txtSup_wk02.Text + "','" + user.lblPrevShipMethod_02.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_02 + "-" + header.v_yr_02 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                      sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_01 + "','" + header.v_wk_01 + "','" + header.v_yr_01 + "','" + user.txtSup_wk01.Text + "','" + user.lblPrevShipMethod_01.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_01 + "-" + header.v_yr_01 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                      sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_0 + "','" + header.v_wk_0 + "','" + header.v_yr_0 + "','" + user.txtSup_wk0.Text + "','" + user.lblPrevShipMethod_0.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_0 + "-" + header.v_yr_0 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";*/
        //                }
        //                else
        //                {
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_1 + "'," + this.wk_num_1 + "," + this.year_1 + "," + new_com_qty_1 + ",'" + cmbSite.SelectedValue + "','" + ship_method1 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_1 + "-" + this.year_1 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty ,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_2 + "'," + this.wk_num_2 + "," + this.year_2 + "," + new_com_qty_2 + ",'" + cmbSite.SelectedValue + "','" + ship_method2 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_2 + "-" + this.year_2 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_3 + "'," + this.wk_num_3 + "," + this.year_3 + "," + new_com_qty_3 + ",'" + cmbSite.SelectedValue + "','" + ship_method3 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_3 + "-" + this.year_3 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_4 + "'," + this.wk_num_4 + "," + this.year_4 + "," + new_com_qty_4 + ",'" + cmbSite.SelectedValue + "','" + ship_method4 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_4 + "-" + this.year_4 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_5 + "'," + this.wk_num_5 + "," + this.year_5 + "," + new_com_qty_5 + ",'" + cmbSite.SelectedValue + "','" + ship_method5 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_5 + "-" + this.year_5 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_6 + "'," + this.wk_num_6 + "," + this.year_6 + "," + new_com_qty_6 + ",'" + cmbSite.SelectedValue + "','" + ship_method6 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_6 + "-" + this.year_6 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_7 + "'," + this.wk_num_7 + "," + this.year_7 + "," + new_com_qty_7 + ",'" + cmbSite.SelectedValue + "','" + ship_method7 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_7 + "-" + this.year_7 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_8 + "'," + this.wk_num_8 + "," + this.year_8 + "," + new_com_qty_8 + ",'" + cmbSite.SelectedValue + "','" + ship_method8 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_8 + "-" + this.year_8 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_9 + "'," + this.wk_num_9 + "," + this.year_9 + "," + new_com_qty_9 + ",'" + cmbSite.SelectedValue + "','" + ship_method9 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_9 + "-" + this.year_9 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_10 + "'," + this.wk_num_10 + "," + this.year_10 + "," + new_com_qty_10 + ",'" + cmbSite.SelectedValue + "','" + ship_method10 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_10 + "-" + this.year_10 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_11 + "'," + this.wk_num_11 + "," + this.year_11 + "," + new_com_qty_11 + ",'" + cmbSite.SelectedValue + "','" + ship_method11 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_11 + "-" + this.year_11 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_12 + "'," + this.wk_num_12 + "," + this.year_12 + "," + new_com_qty_12 + ",'" + cmbSite.SelectedValue + "','" + ship_method12 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_12 + "-" + this.year_12 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_13 + "'," + this.wk_num_13 + "," + this.year_13 + "," + new_com_qty_13 + ",'" + cmbSite.SelectedValue + "','" + ship_method13 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_13 + "-" + this.year_13 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_14 + "'," + this.wk_num_14 + "," + this.year_14 + "," + new_com_qty_14 + ",'" + cmbSite.SelectedValue + "','" + ship_method14 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_14 + "-" + this.year_14 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_15 + "'," + this.wk_num_15 + "," + this.year_15 + "," + new_com_qty_15 + ",'" + cmbSite.SelectedValue + "','" + ship_method15 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_15 + "-" + this.year_15 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_16 + "'," + this.wk_num_16 + "," + this.year_16 + "," + new_com_qty_16 + ",'" + cmbSite.SelectedValue + "','" + ship_method16 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_16 + "-" + this.year_16 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_17 + "'," + this.wk_num_17 + "," + this.year_17 + "," + new_com_qty_17 + ",'" + cmbSite.SelectedValue + "','" + ship_method17 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_17 + "-" + this.year_17 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_18 + "'," + this.wk_num_18 + "," + this.year_18 + "," + new_com_qty_18 + ",'" + cmbSite.SelectedValue + "','" + ship_method18 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_18 + "-" + this.year_18 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_19 + "'," + this.wk_num_19 + "," + this.year_19 + "," + new_com_qty_19 + ",'" + cmbSite.SelectedValue + "','" + ship_method19 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_19 + "-" + this.year_19 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_req_qty,site,ship_method,commit_name,plan_type,po,line) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_20 + "'," + this.wk_num_20 + "," + this.year_20 + "," + new_com_qty_20 + ",'" + cmbSite.SelectedValue + "','" + ship_method20 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + wk_num_20 + "-" + this.year_20 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_04 + "','" + header.v_wk_04 + "','" + header.v_yr_04 + "','" + user.txtSup_wk04.Text + "','" + prev_week_ship_method_4 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_04 + "-" + header.v_yr_04 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_03 + "','" + header.v_wk_03 + "','" + header.v_yr_03 + "','" + user.txtSup_wk03.Text + "','" + prev_week_ship_method_3 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_03 + "-" + header.v_yr_03 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_02 + "','" + header.v_wk_02 + "','" + header.v_yr_02 + "','" + user.txtSup_wk02.Text + "','" + prev_week_ship_method_2 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_02 + "-" + header.v_yr_02 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_01 + "','" + header.v_wk_01 + "','" + header.v_yr_01 + "','" + user.txtSup_wk01.Text + "','" + prev_week_ship_method_1 + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_01 + "-" + header.v_yr_01 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
        //                    sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_0 + "','" + header.v_wk_0 + "','" + header.v_yr_0 + "','" + user.txtSup_wk0.Text + "','" + prev_week_ship_method + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_0 + "-" + header.v_yr_0 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";


        //                }

        //            }
        //            catch
        //            {
        //            }
        //        }
                
        //    }

        //        try
        //        {
           
        //        SqlConnection cn = new SqlConnection(this.constr);
        //        cn.Open();
        //        if (com_id == 0)
        //            sql = sql + "  UPDATE    vc_commit SET  com_id = (SELECT     MAX(auto_id) AS Expr1 FROM  vc_commit AS vc_commit_1) WHERE     (com_id IS NULL)    declare @id as integer SELECT     @id=MAX(auto_id)  FROM  vc_commit select @id ";
        //        //  sql = sql  "  UPDATE    vc_commit SET  com_id = (SELECT     MAX(auto_id) AS Expr1 FROM  vc_commit AS vc_commit_1) WHERE     (com_id IS NULL)    declare @id as integer SELECT     @id=MAX(auto_id)  FROM  vc_commit select @id ";
        //        SqlCommand cmd = new SqlCommand("declare @cnt integer " + sql, cn);
        //        SqlDataAdapter da = new SqlDataAdapter(cmd);
        //        DataSet ds = new DataSet();
        //        da.Fill(ds);
        //        if (com_id == 0)
        //            com_id = Convert.ToInt32(ds.Tables[ds.Tables.Count - 1].Rows[0][0].ToString());

        //        lblCommitId.Text = com_id.ToString();
        //        MessageBox.Show(" New Request has been  commited with commit id " + com_id.ToString());
        //        button3.Enabled = true;
        //        button3.BackColor = System.Drawing.Color.White;
        //        lblStatus.Text = "NEW";
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error :" + ex.Message.ToString());
        //    }


        //}

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //string sql = "";
            //int com_id = 0;
            //string status = "";
            //DataSet ds = getdataSet("select * from vc_commit where com_id='" + Convert.ToInt32(lblCommitId.Text.Replace("COMMIT ID:", "")) + "'");
            //if (ds != null)
            //    if (ds.Tables[0].Rows.Count > 0)
            //        status = ds.Tables[0].Rows[0]["status"].ToString();

            //if (status == "P")
            //{

            //    for (int r = 0; r < distinctValues.Rows.Count; r++)
            //    {
            //        string part = distinctValues.Rows[r]["part_num"].ToString();
            //        // drDemand = dsPart.Tables[0].Select("part_num='" + distinctValues.Rows[r]["part_num"].ToString() + "'");
            //        //   DataRow[] drSupply = dsPart.Tables[1].Select("part_num='" + distinctValues.Rows[r]["part_num"].ToString() + "'");

            //        UserControl1 user = panel2.Controls.Find("dem" + part, true)[0] as UserControl1;
            //        UserControl2 header = this.pnlHeader.Controls.Find("header", true)[0] as UserControl2;


            //        if (user != null)
            //        {


            //            string plan_type = "";
            //            string po = "";
            //            string line = "";

            //            plan_type = user.lblPlanType.Text.ToString();
            //            po = user.lblPONum.Text.ToString();
            //            line = user.lblLineNum.Text.ToString();

            //            int new_com_week_1 = user.p_wk_1;
            //            int new_com_qty_1 = user.newcommit_wk_1;

            //            int new_com_week_2 = user.p_wk_2;
            //            int new_com_qty_2 = user.newcommit_wk_2;

            //            int new_com_week_3 = user.p_wk_3;
            //            int new_com_qty_3 = user.newcommit_wk_3;

            //            int new_com_week_4 = user.p_wk_4;
            //            int new_com_qty_4 = user.newcommit_wk_4;

            //            int new_com_week_5 = user.p_wk_5;
            //            int new_com_qty_5 = user.newcommit_wk_5;

            //            int new_com_week_6 = user.p_wk_6;
            //            int new_com_qty_6 = user.newcommit_wk_6;

            //            int new_com_week_7 = user.p_wk_7;
            //            int new_com_qty_7 = user.newcommit_wk_7;

            //            int new_com_week_8 = user.p_wk_8;
            //            int new_com_qty_8 = user.newcommit_wk_8;

            //            int new_com_week_9 = user.p_wk_9;
            //            int new_com_qty_9 = user.newcommit_wk_9;

            //            int new_com_week_10 = user.p_wk_10;
            //            int new_com_qty_10 = user.newcommit_wk_10;

            //            int new_com_week_11 = user.p_wk_11;
            //            int new_com_qty_11 = user.newcommit_wk_11;

            //            int new_com_week_12 = user.p_wk_12;
            //            int new_com_qty_12 = user.newcommit_wk_12;

            //            int new_com_week_13 = user.p_wk_13;
            //            int new_com_qty_13 = user.newcommit_wk_13;

            //            int new_com_week_14 = user.p_wk_14;
            //            int new_com_qty_14 = user.newcommit_wk_14;

            //            int new_com_week_15 = user.p_wk_15;
            //            int new_com_qty_15 = user.newcommit_wk_15;

            //            int new_com_week_16 = user.p_wk_16;
            //            int new_com_qty_16 = user.newcommit_wk_16;

            //            // int new_com_week_13 = user.p_wk_13;
            //            int new_com_qty_20 = user.newcommit_wk_20;

            //            //  int new_com_week_14 = user.p_wk_14;
            //            int new_com_qty_19 = user.newcommit_wk_19;

            //            //  int new_com_week_15 = user.p_wk_15;
            //            int new_com_qty_18 = user.newcommit_wk_18;

            //            //  int new_com_week_16 = user.p_wk_16;
            //            int new_com_qty_17 = user.newcommit_wk_17;

            //            int vendor_qty_1 = 0;
            //            int vendor_qty_2 = 0;
            //            int vendor_qty_3 = 0;
            //            int vendor_qty_4 = 0;
            //            int vendor_qty_5 = 0;
            //            int vendor_qty_6 = 0;
            //            int vendor_qty_7 = 0;
            //            int vendor_qty_8 = 0;
            //            int vendor_qty_9 = 0;
            //            int vendor_qty_10 = 0;
            //            int vendor_qty_11 = 0;
            //            int vendor_qty_12 = 0;
            //            int vendor_qty_13 = 0;
            //            int vendor_qty_14 = 0;
            //            int vendor_qty_15 = 0;
            //            int vendor_qty_16 = 0;
            //            int vendor_qty_17 = 0;
            //            int vendor_qty_18 = 0;
            //            int vendor_qty_19 = 0;
            //            int vendor_qty_20 = 0;
            //            if (user.txtVendorQty1.Text != "")
            //            {
            //                vendor_qty_1 = Convert.ToInt32(user.txtVendorQty1.Text);
            //            }
            //            if (user.txtVendorQty2.Text != "")
            //                vendor_qty_2 = Convert.ToInt32(user.txtVendorQty2.Text);
            //            if (user.txtVendorQty3.Text != "")
            //                vendor_qty_3 = Convert.ToInt32(user.txtVendorQty3.Text);

            //            if (user.txtVendorQty4.Text != "")
            //                vendor_qty_4 = Convert.ToInt32(user.txtVendorQty4.Text);

            //            if (user.txtVendorQty5.Text != "")
            //                vendor_qty_5 = Convert.ToInt32(user.txtVendorQty5.Text);

            //            if (user.txtVendorQty6.Text != "")
            //                vendor_qty_6 = Convert.ToInt32(user.txtVendorQty6.Text);

            //            if (user.txtVendorQty7.Text != "")
            //                vendor_qty_7 = Convert.ToInt32(user.txtVendorQty7.Text);

            //            if (user.txtVendorQty8.Text != "")
            //                vendor_qty_8 = Convert.ToInt32(user.txtVendorQty8.Text);

            //            if (user.txtVendorQty9.Text != "")
            //                vendor_qty_9 = Convert.ToInt32(user.txtVendorQty9.Text);

            //            if (user.txtVendorQty10.Text != "")
            //                vendor_qty_10 = Convert.ToInt32(user.txtVendorQty10.Text);

            //            if (user.txtVendorQty11.Text != "")
            //                vendor_qty_11 = Convert.ToInt32(user.txtVendorQty11.Text);

            //            if (user.txtVendorQty12.Text != "")
            //                vendor_qty_12 = Convert.ToInt32(user.txtVendorQty12.Text);

            //            if (user.txtVendorQty13.Text != "")
            //                vendor_qty_13 = Convert.ToInt32(user.txtVendorQty13.Text);


            //            if (user.txtVendorQty14.Text != "")
            //                vendor_qty_14 = Convert.ToInt32(user.txtVendorQty14.Text);

            //            if (user.txtVendorQty15.Text != "")
            //                vendor_qty_15 = Convert.ToInt32(user.txtVendorQty15.Text);

            //            if (user.txtVendorQty16.Text != "")
            //                vendor_qty_16 = Convert.ToInt32(user.txtVendorQty16.Text);

            //            if (user.txtVendorQty17.Text != "")
            //                vendor_qty_17 = Convert.ToInt32(user.txtVendorQty17.Text);

            //            if (user.txtVendorQty18.Text != "")

            //                vendor_qty_18 = Convert.ToInt32(user.txtVendorQty18.Text);

            //            if (user.txtVendorQty19.Text != "")
            //                vendor_qty_19 = Convert.ToInt32(user.txtVendorQty19.Text);

            //            if (user.txtVendorQty20.Text != "")
            //                vendor_qty_20 = Convert.ToInt32(user.txtVendorQty20.Text);

            //            int prev_week_supply = 0;
            //            int prev_week_supply_1 = 0;
            //            int prev_week_supply_2 = 0;
            //            int prev_week_supply_3 = 0;
            //            int prev_week_supply_4 = 0;

            //            if (user.txtSup_wk0.Text != "")
            //                prev_week_supply = Convert.ToInt32(user.txtSup_wk0.Text);
            //            if (user.txtSup_wk01.Text != "")
            //                prev_week_supply_1 = Convert.ToInt32(user.txtSup_wk01.Text);
            //            if (user.txtSup_wk02.Text != "")
            //                prev_week_supply_2 = Convert.ToInt32(user.txtSup_wk02.Text);
            //            if (user.txtSup_wk03.Text != "")
            //                prev_week_supply_3 = Convert.ToInt32(user.txtSup_wk03.Text);
            //            if (user.txtSup_wk04.Text != "")
            //                prev_week_supply_4 = Convert.ToInt32(user.txtSup_wk04.Text);


            //            string prev_week_ship_method = "A";
            //            string prev_week_ship_method_1 = "A";
            //            string prev_week_ship_method_2 = "A";
            //            string prev_week_ship_method_3 = "A";
            //            string prev_week_ship_method_4 = "A";
            //            if (user.radioOceanWk_04.Checked)
            //                prev_week_ship_method_4 = "O";
            //            if (user.radioOceanWk_03.Checked)
            //                prev_week_ship_method_3 = "O";
            //            if (user.radioOceanWk_02.Checked)
            //                prev_week_ship_method_2 = "O";
            //            if (user.radioOceanWk_01.Checked)
            //                prev_week_ship_method_1 = "O";
            //            if (user.radioOceanWk_0.Checked)
            //                prev_week_ship_method = "O";



            //            if (lblCommitId.Text.Replace("COMMIT ID:", "") != "")
            //            {
            //                com_id = Convert.ToInt32(lblCommitId.Text.Replace("COMMIT ID:", ""));
            //            }
            //            if (com_id != 0)
            //            {

            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_1 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_1;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_2 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_2;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_3 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_3;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_4 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_4;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_5 + "' , status='C',commit_date=getdate()  where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_5;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_6 + "'  , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_6;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_7 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_7;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_8 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_8;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_9 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_9;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_10 + "', status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_10;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_11 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_11;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_12 + "', status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_12;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_13 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_13;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_14 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_14;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_15 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_15;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_16 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_16;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_17 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_17;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_18 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_18;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_19 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_19;
            //                sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_20 + "' , status='C',commit_date=getdate() where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_20;

            //                sql = sql + "  update vc_commit set new_com_qty=prev_com_qty,ship_method=prev_ship_method,status='C',commit_date=getdate()  where   com_id=" + com_id + " and part_num='" + part + "' and week < " +  this.wk_num_1;
            //                //*****************prev_commit data for week03 ******************************************************************
            //                // sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_03 + "  if (@cnt=0 ) insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_03 + "','" + header.v_wk_03 + "','" + header.v_yr_03 + "','" + user.txtSup_wk03.Text + "','" + user.lblPrevShipMethod_03.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_03 + "-" + header.v_yr_03 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "') else ";
            //                //sql = sql + "  update vc_commit set new_com_qty=prev_com_qty,ship_method=prev_ship_method,status='C',commit_date=getdate()  where  com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_03;
            //                ////*****************prev_commit data for week02 ******************************************************************
            //                //// sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_02 + "  if (@cnt=0 ) insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_02 + "','" + header.v_wk_02 + "','" + header.v_yr_02 + "','" + user.txtSup_wk02.Text + "','" + user.lblPrevShipMethod_02.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_02 + "-" + header.v_yr_02 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "') else ";
            //                //sql = sql + "  update vc_commit set new_com_qty=prev_com_qty,ship_method=prev_ship_method,status='C',commit_date=getdate()  where com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_02;
            //                ////*****************prev_commit data for week01 ******************************************************************
            //                ////  sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_01 + "  if (@cnt=0 ) insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_01 + "','" + header.v_wk_01 + "','" + header.v_yr_01 + "','" + user.txtSup_wk01.Text + "','" + user.lblPrevShipMethod_01.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_01 + "-" + header.v_yr_01 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "') else ";
            //                //sql = sql + "  update vc_commit set new_com_qty=prev_com_qty,ship_method=prev_ship_method,status='C',commit_date=getdate()  where  com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_01;
            //                ////*****************prev_commit data for week0 ******************************************************************
            //                //// sql = sql + "  select @cnt=count(*) from vc_commit where status='N' and com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_0 + "  if (@cnt=0 ) insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_0 + "','" + header.v_wk_0 + "','" + header.v_yr_0 + "','" + user.txtSup_wk0.Text + "','" + user.lblPrevShipMethod_0.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_0 + "-" + header.v_yr_0 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "') else ";
            //                //sql = sql + "  update vc_commit set new_com_qty=prev_com_qty,ship_method=prev_ship_method,status='C',commit_date=getdate()  where com_id=" + com_id + " and part_num='" + part + "' and week=" + header.v_wk_0;

            //                /*  sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_04 + "','" + header.v_wk_04 + "','" + header.v_yr_04 + "','" + user.txtSup_wk04.Text + "','" + user.lblPrevShipMethod_04.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_04 + "-" + header.v_yr_04 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
            //                sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_03 + "','" + header.v_wk_03 + "','" + header.v_yr_03 + "','" + user.txtSup_wk03.Text + "','" + user.lblPrevShipMethod_03.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_03 + "-" + header.v_yr_03 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
            //                sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_02 + "','" + header.v_wk_02 + "','" + header.v_yr_02 + "','" + user.txtSup_wk02.Text + "','" + user.lblPrevShipMethod_02.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_02 + "-" + header.v_yr_02 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
            //                sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_01 + "','" + header.v_wk_01 + "','" + header.v_yr_01 + "','" + user.txtSup_wk01.Text + "','" + user.lblPrevShipMethod_01.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_01 + "-" + header.v_yr_01 + "','" + plan_type.Substring(0, 1) + "','" + po + "','" + line + "')";
            //                sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) values('N','" + this.cmbSite.SelectedValue + "','" + this.cmbVendor.SelectedValue + "',getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),'" + part + "','" + header.v_date_0 + "','" + header.v_wk_0 + "','" + header.v_yr_0 + "','" + user.txtSup_wk0.Text + "','" + user.lblPrevShipMethod_0.Text.ToString().Substring(0, 1) + "','" + this.cmbSite.SelectedValue + "." + cmbVendor.Text + "." + part + "." + header.v_wk_0 + "-" + header.v_yr_0 + "','" + plan_type.Substring(0,1) + "','" + po + "','" + line + "')";*/
            //            }

            //        }
            //    }
            //    try
            //    {
            //        SqlConnection cn = new SqlConnection(this.constr);
            //        cn.Open();
            //        sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) select 'N',site,vend_id,getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),part_num,date,week,year,new_com_qty,ship_method,commit_name,plan_type,po,line from vc_commit where com_id=" + com_id + " ";
            //        sql = sql + "  update vc_commit set com_id =(select max(auto_id) from vc_commit) where com_id is null";
            //        SqlCommand cmd = new SqlCommand(sql, cn);
            //        cmd.ExecuteNonQuery();
            //        button2.Enabled = false;
            //        button2.BackColor = System.Drawing.Color.IndianRed;

            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Error :" + ex.Message.ToString());
            //    }
            //    MessageBox.Show(" Vendor Commit has been  commited for " + lblCommitId.Text);
            //    //getPart();
            //}
            //else
            //{
            //    MessageBox.Show("Vendor commit can not be saved. Please check the status for this commit");
            //}


        }

        //private void button3_Click(object sender, EventArgs e)
        //{

        //    /*  string sql = "";
        //      for (int r = 0; r < distinctValues.Rows.Count; r++)
        //      {
        //          string part = distinctValues.Rows[r]["part_num"].ToString();
        //          // drDemand = dsPart.Tables[0].Select("part_num='" + distinctValues.Rows[r]["part_num"].ToString() + "'");
        //          // DataRow[] drSupply = dsPart.Tables[1].Select("part_num='" + distinctValues.Rows[r]["part_num"].ToString() + "'");
        //          UserControl1 user = panel2.Controls.Find("dem" + part, true)[0] as UserControl1;
        //          if (user != null)
        //          {
        //              int new_com_week_1 = user.p_wk_1;
        //              int new_com_qty_1 = 0;
        //              int new_com_qty_2 = 0;
        //              int new_com_qty_3 = 0;
        //              int new_com_qty_4 = 0;
        //              int new_com_qty_5 = 0;
        //              int new_com_qty_6 = 0;
        //              int new_com_qty_7 = 0;
        //              int new_com_qty_8 = 0;
        //              int new_com_qty_9 = 0;
        //              int new_com_qty_10 = 0;
        //              int new_com_qty_11 = 0;
        //              int new_com_qty_12 = 0;
        //              int new_com_qty_13= 0;
        //              int new_com_qty_14= 0;
        //              int new_com_qty_15= 0;
        //              int new_com_qty_16= 0;
        //              int new_com_qty_17= 0;
        //              int new_com_qty_18= 0;
        //              int new_com_qty_19= 0;
        //              int new_com_qty_20 = 0;
        //              if (user.txtVendorQty1.Text != "")
        //                   new_com_qty_1 = Convert.ToInt32(user.txtVendorQty1.Text);
                    
        //              string ship_method1 = user.shipMethod[0].ToString();

        //              int new_com_week_2 = user.p_wk_2;
        //              if (user.txtVendorQty2.Text.ToString().Trim() != "")
        //                    new_com_qty_2 = Convert.ToInt32(user.txtVendorQty2.Text);
        //              string ship_method2 = user.shipMethod[1].ToString();

        //              int new_com_week_3 = user.p_wk_3;
        //              if (user.txtVendorQty3.Text.ToString().Trim() != "")
        //              {
        //                   new_com_qty_3 = Convert.ToInt32(user.txtVendorQty3.Text);
        //              }
        //              string ship_method3 = user.shipMethod[2].ToString();

        //              int new_com_week_4 = user.p_wk_4;
        //              if (user.txtVendorQty4.Text.ToString().Trim() != "")
                      
        //                 new_com_qty_4 = Convert.ToInt32(user.txtVendorQty4.Text);
                      
        //              string ship_method4 = user.shipMethod[3].ToString();

        //              int new_com_week_5 = user.p_wk_5;
        //              if (user.txtVendorQty5.Text.ToString().Trim() != "")
                   
        //              new_com_qty_5 = Convert.ToInt32(user.txtVendorQty5.Text);
        //              string ship_method5 = user.shipMethod[4].ToString();

        //              int new_com_week_6 = user.p_wk_6;
        //              if (user.txtVendorQty6.Text.ToString().Trim() != "")
        //                   new_com_qty_6 = Convert.ToInt32(user.txtVendorQty6.Text);
        //              string ship_method6 = user.shipMethod[5].ToString();

        //              int new_com_week_7 = user.p_wk_7;
        //              if (user.txtVendorQty7.Text.ToString().Trim() != "")
        //                new_com_qty_7 = Convert.ToInt32(user.txtVendorQty7.Text);
        //              string ship_method7 = user.shipMethod[6].ToString();

        //              int new_com_week_8 = user.p_wk_8;
        //              if (user.txtVendorQty8.Text.ToString().Trim() != "")                      
        //                 new_com_qty_8 = Convert.ToInt32(user.txtVendorQty8.Text);

        //              string ship_method8 = user.shipMethod[7].ToString();

        //              int new_com_week_9 = user.p_wk_9;
        //              if (user.txtVendorQty9.Text.ToString().Trim() != "")
        //               new_com_qty_9 = Convert.ToInt32(user.txtVendorQty9.Text);
        //              string ship_method9 = user.shipMethod[8].ToString();

        //              int new_com_week_10 = user.p_wk_10;
        //              if (user.txtVendorQty10.Text.ToString().Trim() != "")
        //              new_com_qty_10 = Convert.ToInt32(user.txtVendorQty10.Text);
        //              string ship_method10 = user.shipMethod[9].ToString();

        //              int new_com_week_11 = user.p_wk_11;
        //              if (user.txtVendorQty11.Text.ToString().Trim() != "")
        //              new_com_qty_11 = Convert.ToInt32(user.txtVendorQty11.Text);
        //              string ship_method11 = user.shipMethod[10].ToString();

        //              int new_com_week_12 = user.p_wk_12;
        //              if (user.txtVendorQty12.Text.ToString().Trim() != "")
        //               new_com_qty_12 = Convert.ToInt32(user.txtVendorQty12.Text);
        //              string ship_method12 = user.shipMethod[11].ToString();

        //              int new_com_week_13 = user.p_wk_13;
        //              if (user.txtVendorQty13.Text.ToString().Trim() != "")
        //               new_com_qty_13 = Convert.ToInt32(user.txtVendorQty13.Text);
        //              string ship_method13 = user.shipMethod[12].ToString();

        //              int new_com_week_14 = user.p_wk_14;
        //              if (user.txtVendorQty14.Text.ToString().Trim() != "")
        //                new_com_qty_14 = Convert.ToInt32(user.txtVendorQty14.Text);
        //              string ship_method14 = user.shipMethod[13].ToString();

        //              int new_com_week_15 = user.p_wk_15;
        //              if (user.txtVendorQty15.Text.ToString().Trim() != "")                      
        //                  new_com_qty_15 = Convert.ToInt32(user.txtVendorQty15.Text);
        //              string ship_method15 = user.shipMethod[14].ToString();

        //              int new_com_week_16 = user.p_wk_16;
        //              if (user.txtVendorQty16.Text.ToString().Trim() != "")                     
        //                  new_com_qty_16 = Convert.ToInt32(user.txtVendorQty16.Text);
        //              string ship_method16 = user.shipMethod[15].ToString();


        //              int new_com_week_17 = user.p_wk_17;
        //              if (user.txtVendorQty17.Text.ToString().Trim() != "")
        //                        new_com_qty_17 = Convert.ToInt32(user.txtVendorQty17.Text);
                      
        //              string ship_method17 = user.shipMethod[16].ToString();

        //              int new_com_week_18 = user.p_wk_18;
        //              if (user.txtVendorQty18.Text.ToString().Trim() != "")
        //                        new_com_qty_18 = Convert.ToInt32(user.txtVendorQty18.Text);
        //              string ship_method18 = user.shipMethod[17].ToString();

        //              int new_com_week_19 = user.p_wk_19;
        //              if (user.txtVendorQty19.Text.ToString().Trim() != "")
                    
        //                new_com_qty_19 = Convert.ToInt32(user.txtVendorQty19.Text);
        //              string ship_method19 = user.shipMethod[18].ToString();

        //              int new_com_week_20 = user.p_wk_20;
        //              if (user.txtVendorQty20.Text.ToString().Trim() != "")
        //                   new_com_qty_20 = Convert.ToInt32(user.txtVendorQty20.Text);
                      
        //              string ship_method20 = user.shipMethod[19].ToString();



        //              //int vendor_qty_1 = user.newcommit_wk_1;
        //              //int vendor_qty_2 = user.newcommit_wk_2;
        //              //int vendor_qty_3 = user.newcommit_wk_3;
        //              //int vendor_qty_4 = user.newcommit_wk_4;
        //              //int vendor_qty_5 = user.newcommit_wk_5;
        //              //int vendor_qty_6 = user.newcommit_wk_6;
        //              //int vendor_qty_7 = user.newcommit_wk_7;               
        //              //int vendor_qty_8 = user.newcommit_wk_8;
        //              //int vendor_qty_9 = user.newcommit_wk_9;
        //              //int vendor_qty_10 = user.newcommit_wk_10;
        //              //int vendor_qty_11 = user.newcommit_wk_11;
        //              //int vendor_qty_12 = user.newcommit_wk_12;
        //              //int vendor_qty_13 = user.newcommit_wk_13;
        //              //int vendor_qty_14 = user.newcommit_wk_14;
        //              //int vendor_qty_15 = user.newcommit_wk_15;
        //              //int vendor_qty_16 = user.newcommit_wk_16;
        //              //int vendor_qty_20 = user.newcommit_wk_20;
        //              //int vendor_qty_18 = user.newcommit_wk_18;
        //              //int vendor_qty_17 = user.newcommit_wk_17;
        //              //int vendor_qty_19 = user.newcommit_wk_19;

        //              // int prev_qty_1 = 0;
        //              // int prev_qty_2 = 0;
        //              // int prev_qty_3 = 0;
        //              // int prev_qty_4 = 0;
        //              // int prev_qty_5 = 0;
        //              // int prev_qty_6 = 0;
        //              // int prev_qty_7 = 0;
        //              // int prev_qty_8 = 0;
        //              // int prev_qty_9 = 0;
        //              // int prev_qty_10 = 0;
        //              // int prev_qty_11 = 0;
        //              // int prev_qty_12 = 0;
        //              // int prev_qty_13 = 0;
        //              // int prev_qty_14 = 0;
        //              // int prev_qty_15 = 0;
        //              // int prev_qty_16 = 0;
        //              // int prev_qty_17 = 0;
        //              // int prev_qty_18 = 0;
        //              // int prev_qty_19 = 0;
        //              //int prev_qty_20= 0;
        //              //for (int r1 = 0; r1 < drSupply.Length; r1++)
        //              //{
        //              //    if (drSupply[r1]["part_num"].ToString() == part)
        //              //    {
        //              //        prev_qty_1 = Convert.ToInt32(drSupply[r1][0].ToString());
        //              //        prev_qty_2 = Convert.ToInt32(drSupply[r1][1].ToString());
        //              //        prev_qty_3 = Convert.ToInt32(drSupply[r1][2].ToString());
        //              //        prev_qty_4 = Convert.ToInt32(drSupply[r1][3].ToString());
        //              //        prev_qty_5 = Convert.ToInt32(drSupply[r1][4].ToString());
        //              //        prev_qty_6 = Convert.ToInt32(drSupply[r1][5].ToString());
        //              //        prev_qty_7 = Convert.ToInt32(drSupply[r1][6].ToString());
        //              //        prev_qty_8 = Convert.ToInt32(drSupply[r1][7].ToString());
        //              //        prev_qty_9 = Convert.ToInt32(drSupply[r1][8].ToString());
        //              //        prev_qty_10 = Convert.ToInt32(drSupply[r1][9].ToString());
        //              //        prev_qty_11 = Convert.ToInt32(drSupply[r1][10].ToString());
        //              //        prev_qty_12 = Convert.ToInt32(drSupply[r1][11].ToString());
        //              //        prev_qty_13 = Convert.ToInt32(drSupply[r1][12].ToString());
        //              //        prev_qty_14 = Convert.ToInt32(drSupply[r1][13].ToString());
        //              //        prev_qty_15 = Convert.ToInt32(drSupply[r1][14].ToString());
        //              //        prev_qty_16 = Convert.ToInt32(drSupply[r1][15].ToString());
        //              //        prev_qty_17 = Convert.ToInt32(drSupply[r1][16].ToString());
        //              //        prev_qty_18 = Convert.ToInt32(drSupply[r1][17].ToString());
        //              //        prev_qty_19 = Convert.ToInt32(drSupply[r1][18].ToString());
        //              //        prev_qty_20 = Convert.ToInt32(drSupply[r1][19].ToString());
        //              //    }
        //              //}
        //              //prev_qty_20 = Convert.ToInt32(drSupply[r1][20].ToString());
        //              int com_id = 0;
                     
        //              if (com_id != 0)
        //              {

        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_1 + "',ship_method='" + ship_method1 + "'  where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_1;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_2 + "',ship_method='" + ship_method2 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_2;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_3 + "',ship_method='" + ship_method3 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_3;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_4 + "',ship_method='" + ship_method4 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_4;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_5 + "',ship_method='" + ship_method5 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_5;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_6 + "',ship_method='" + ship_method6 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_6;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_7 + "',ship_method='" + ship_method7 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_7;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_8 + "',ship_method='" + ship_method8 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_8;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_9 + "',ship_method='" + ship_method9 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_9;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_10 + "',ship_method='" + ship_method10 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_10;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_11 + "',ship_method='" + ship_method11 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_11;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_12 + "',ship_method='" + ship_method12 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_12;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_13 + "',ship_method='" + ship_method13 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_13;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_14 + "',ship_method='" + ship_method14 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_14;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_15 + "',ship_method='" + ship_method15 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_15;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_16 + "',ship_method='" + ship_method16 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_16;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_17 + "',ship_method='" + ship_method17 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_17;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_18 + "',ship_method='" + ship_method18 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_18;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_19 + "',ship_method='" + ship_method19 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_19;
        //                  sql = sql + "  update vc_commit set new_com_qty='" + new_com_qty_20 + "',ship_method='" + ship_method20 + "' where  com_id=" + com_id + " and part_num='" + part + "' and week=" + this.wk_num_20;
        //              }
        //              else
        //              {
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty ,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_1 + "'," + this.wk_num_1 + "," + this.year_1 + "," + new_com_qty_1 + ",'" + cmbSite.SelectedValue + "','" + ship_method1 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty ,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_2 + "'," + this.wk_num_2 + "," + this.year_2 + "," + new_com_qty_2 + ",'" + cmbSite.SelectedValue + "','" + ship_method2 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_3 + "'," + this.wk_num_3 + "," + this.year_3 + "," + new_com_qty_3 + ",'" + cmbSite.SelectedValue + "','" + ship_method3 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_4 + "'," + this.wk_num_4 + "," + this.year_4 + "," + new_com_qty_4 + ",'" + cmbSite.SelectedValue + "','" + ship_method4 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_5 + "'," + this.wk_num_5 + "," + this.year_5 + "," + new_com_qty_5 + ",'" + cmbSite.SelectedValue + "','" + ship_method5 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_6 + "'," + this.wk_num_6 + "," + this.year_6 + "," + new_com_qty_6 + ",'" + cmbSite.SelectedValue + "','" + ship_method6 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_7 + "'," + this.wk_num_7 + "," + this.year_7 + "," + new_com_qty_7 + ",'" + cmbSite.SelectedValue + "','" + ship_method7 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_8 + "'," + this.wk_num_8 + "," + this.year_8 + "," + new_com_qty_8 + ",'" + cmbSite.SelectedValue + "','" + ship_method8 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_9 + "'," + this.wk_num_9 + "," + this.year_9 + "," + new_com_qty_9 + ",'" + cmbSite.SelectedValue + "','" + ship_method9 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_10 + "'," + this.wk_num_10 + "," + this.year_10 + "," + new_com_qty_10 + ",'" + cmbSite.SelectedValue + "','" + ship_method10 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_11 + "'," + this.wk_num_11 + "," + this.year_11 + "," + new_com_qty_11 + ",'" + cmbSite.SelectedValue + "','" + ship_method11 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_12 + "'," + this.wk_num_12 + "," + this.year_12 + "," + new_com_qty_12 + ",'" + cmbSite.SelectedValue + "','" + ship_method12 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_13 + "'," + this.wk_num_13 + "," + this.year_13 + "," + new_com_qty_13 + ",'" + cmbSite.SelectedValue + "','" + ship_method13 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_14 + "'," + this.wk_num_14 + "," + this.year_14 + "," + new_com_qty_14 + ",'" + cmbSite.SelectedValue + "','" + ship_method14 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_15 + "'," + this.wk_num_15 + "," + this.year_15 + "," + new_com_qty_15 + ",'" + cmbSite.SelectedValue + "','" + ship_method15 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_16 + "'," + this.wk_num_16 + "," + this.year_16 + "," + new_com_qty_16 + ",'" + cmbSite.SelectedValue + "','" + ship_method16 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_17 + "'," + this.wk_num_17 + "," + this.year_17 + "," + new_com_qty_17 + ",'" + cmbSite.SelectedValue + "','" + ship_method17 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_18 + "'," + this.wk_num_18 + "," + this.year_18 + "," + new_com_qty_18 + ",'" + cmbSite.SelectedValue + "','" + ship_method18 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_19 + "'," + this.wk_num_19 + "," + this.year_19 + "," + new_com_qty_19 + ",'" + cmbSite.SelectedValue + "','" + ship_method19 + "')";
        //                  sql = sql + "  insert into vc_commit (vend_id,status,create_date,create_week,create_year,part_num,date,week,year,new_com_qty,site,ship_method) values (" + cmbVendor.SelectedValue + ",'N',getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate(getdate()),'" + part + "','" + this.date_20 + "'," + this.wk_num_20 + "," + this.year_20 + "," + new_com_qty_20 + ",'" + cmbSite.SelectedValue + "','" + ship_method20 + "')";


        //              }

        //          }

        //      }*/
        //    int com_id = 0;
        //    if (lblCommitId.Text.Replace(" ", "").Trim() != "")
        //    {
        //        com_id = Convert.ToInt32(lblCommitId.Text.Replace("COMMIT ID:", "").Trim());

        //        try
        //        {
        //            SqlConnection cn = new SqlConnection(this.constr);
        //            cn.Open();

        //            string sql = "  UPDATE    vc_commit SET  status='P', ready_to_send_date=getdate()  where status='N' and site='" + cmbSite.SelectedValue.ToString() + "' and com_id=" + com_id.ToString();


        //            SqlCommand cmd = new SqlCommand(sql, cn);
        //            SqlDataAdapter da = new SqlDataAdapter(cmd);
        //            DataSet ds = new DataSet();
        //            da.Fill(ds);

        //            lblCommitId.Text = com_id.ToString();
        //            MessageBox.Show(" Status of new Request has been updated to READY TO SEND for " + com_id.ToString());
        //            lblStatus.Text = "READY TO SEND";
        //            button1.Enabled = false;
        //            button1.BackColor = System.Drawing.Color.IndianRed;
        //            button3.Enabled = false;
        //            button3.BackColor = System.Drawing.Color.IndianRed;
        //            button2.Enabled = true;
        //            button2.BackColor = System.Drawing.Color.White;
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Error :" + ex.Message.ToString());
        //        }

        //    }


        //    /*
        //    int com_id = 0;
        //    if (lblCommitId.Text.Replace("COMMIT ID :", "") != "")
        //    {
        //        com_id = Convert.ToInt32(lblCommitId.Text.Replace("COMMIT ID :", ""));
             

        //    string sql = "update vc_commit set status='P' where  com_id=" + com_id  ;
        //    try
        //    {
        //        SqlConnection cn = new SqlConnection(this.constr);
        //        cn.Open();
        //        SqlCommand cmd = new SqlCommand(sql, cn);
        //        cmd.ExecuteNonQuery();

        //        MessageBox.Show(" Requested quantity has been updated to send Vendor ");
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error :" + ex.Message.ToString());
        //    }
        //    }*/

        //}

        //public string CommitNewRequirement( string part_num, int prev_com_week, int prev_com_year, string new_com_date, int new_com_qty,int new_com_week, int new_com_year,string  site )
        //{
        //     //string sqlInsert="insert into vc_commit (ven_id,create_date,create_week,create_year,part_num,prev_com_date,prev_com_week,prev_com_year,new_com_date,new_com_qty,new_com_week,new_com_year,site) values (" + +",getdate(),dbo.udf_GetISOWeekNumberFromDate(getdate()),dbo.udf_GetISOYearFromDate((getdate()),'" +part_num+"',"+prev_com_week+","+ prev_com_year+ ",'" +new_com_date+"',"+ new_com_qty+","+new_com_week+ "," + new_com_year+ ",'"+ site+ "')";
        //   //  return sqlInsert;

        //}
        public void loadVendorMaster()
        {
            grdVendor.DataSource = null;

            try
            {
                string sqlGrid = " SELECT [OptionPN], [ASIA], [EUROPE], [NORTH AMERICA], [SOUTH AMERICA], [TAIWAN]  FROM [MIMDIST].[dbo].[MRP_C_RegionOption] ";
                grdVendor.DataSource = getdataSet(sqlGrid).Tables[0];

            }
            catch
            {

            }
       

        }

        private void delete_Click(object sender, EventArgs e)
        {
            // sender.name
        }

        public void loadVendorPartMaster()
        {
            string strVendPart = " SELECT     vend_part_id, vend_id, part_num AS [PART NUM], short_desc AS [PART DESC], std_ship_meth AS [SHIP METHOD], lead_time_air AS [LEAD TIME AIR], lead_time_ocean AS [LEAD TIME OCEAN], lead_time_ground AS [LEAD TIME GROUND], plan_type AS [PLAN TYPE], curr_po_at [CURR PO AT], curr_line_at [CURR LINE AT], curr_po_nl as  [CURR PO NL], curr_line_at AS [CURR LINE NL], curr_po_sg as  [CURR PO SG], curr_line_sg AS [CURR LINE SG], std_box_qty as [STD BOX QTY], std_pallet_qty as [STD PALLET QTY], ord_multiple as [ORD MULTIPLE], min_ord_qty [MIM ORD QTY], factory_fob as [FACTORY FOB], CASE WHEN active='Y' THEN 'ACTIVE' ELSE 'INACTIVE' END AS ACTIVE  FROM   vc_vend_part WHERE   vend_id = " + cmbVendPart.SelectedValue;
            if (cmbStatus.SelectedIndex != 0)
                strVendPart = strVendPart + " and active='" + cmbStatus.SelectedValue + "'";
            strVendPart = strVendPart + "  Order by Part_num ";
            grdVendPart.DataSource = null;
            try
            {
                for (int i = 0; i <= grdVendPart.Columns.Count; i++)
                    grdVendPart.Columns.RemoveAt(0);
            }
            catch
            {
            }


            try
            {
                DataSet dsVendPart = getdataSet(strVendPart);
                grdVendPart.DataSource = dsVendPart.Tables[0];

                DataGridViewTextBoxColumn grdVendPartEdit = new DataGridViewTextBoxColumn();
                grdVendPartEdit.Name = "btnEdit";
                grdVendPartEdit.HeaderText = "EDIT";
                grdVendPart.Columns.Add(grdVendPartEdit);
                //DataGridViewCheckBoxColumn grdVendPartDelete = new DataGridViewCheckBoxColumn();
                //grdVendPartDelete.Name = "chkDelete";
                //grdVendPartDelete.HeaderText = "DELETE";
                //grdVendPart.Columns.Add(grdVendPartDelete);
                grdVendPart.Columns[0].Visible = false;
                grdVendPart.Columns[1].Visible = false;
                grdVendPart.Columns[2].Width = 100;
                grdVendPart.Columns[3].Width = 140;
                grdVendPart.Columns[4].Width = 80;
                grdVendPart.Columns[5].Width = 60;
                grdVendPart.Columns[6].Width = 60;
                grdVendPart.Columns[7].Width = 60;
                grdVendPart.Columns[8].Width = 60;
                grdVendPart.Columns[9].Width = 70;
                grdVendPart.Columns[10].Width = 60;
                grdVendPart.Columns[11].Width = 70;
                grdVendPart.Columns[12].Width = 60;
                grdVendPart.Columns[13].Width = 70;
                grdVendPart.Columns[14].Width = 60;
                grdVendPart.Columns[15].Width = 60;
                grdVendPart.Columns[16].Width = 60;
                grdVendPart.Columns[17].Width = 65;
                grdVendPart.Columns[18].Width = 60;
                grdVendPart.Columns[19].Width = 65;
                grdVendPart.Columns[20].Width = 65;
                grdVendPart.Columns[21].Width = 65;


            }
            catch (Exception e)
            {
                MessageBox.Show("Error :" + e.Message.ToString());
            }

        }

        private void tabControl1_Click(object sender, EventArgs e)
        {

            //if (tabControl1.SelectedIndex == 2)
            //{
            //    loadVendorFileCreate();

            //}
            if (tabControl1.SelectedIndex ==0)
            {
               // loadVendor();
                loadVendorMaster();
            }
            if (tabControl1.SelectedIndex ==1)
            {
                //loadVendor();
                loadComboVendor();
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void grdVendor_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
           

            if (e.RowIndex > -1 && e.ColumnIndex == 4)
            {

                DataGridViewTextBoxCell buttonCell = grdVendor.Rows[e.RowIndex].Cells[4] as DataGridViewTextBoxCell;
                buttonCell.Value = "DELETE";
                buttonCell.Style.ForeColor = Color.Red;

            }
            if (e.RowIndex > -1 && e.ColumnIndex == 3)
            {

                DataGridViewTextBoxCell buttonCell = grdVendor.Rows[e.RowIndex].Cells[3] as DataGridViewTextBoxCell;
                buttonCell.Value = "EDIT";
                buttonCell.Style.ForeColor = Color.Blue;

            }

        }

        private void grdVendor_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4)
            {
                int auto_id = Convert.ToInt32(this.grdVendor.Rows[e.RowIndex].Cells[0].Value);
                deleteRow(auto_id);
            }
            if (e.ColumnIndex == 3)
            {
                lblAutoID.Text = this.grdVendor.Rows[e.RowIndex].Cells[0].Value.ToString();
                txtVendorCode.Text = this.grdVendor.Rows[e.RowIndex].Cells[1].Value.ToString();
                txtVendorType.Text = this.grdVendor.Rows[e.RowIndex].Cells[2].Value.ToString();
            }

        }
        public void deleteRow(int id)
        {
            string sqlQry = "update vc_vend_master set active='N' where vend_id=" + id;
            try
            {
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlCommand cmd = new SqlCommand(sqlQry, cn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Vendor has  been deleted Successfully");
                loadVendorMaster();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Vendor has not been deleted due to some error. Please check");

            }
        }

        public void deleteVendPartRow(int id)
        {
            string sqlQry = "update   vc_vend_part set active='N' where vend_part_id=" + id;
            try
            {
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlCommand cmd = new SqlCommand(sqlQry, cn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Part has been deleted Successfully");
                loadVendorPartMaster();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Part has not been deleted due to some error. Please check");

            }

        }


        public void clearAllVendor()
        {
            txtVendorCode.Text = "";
            txtVendorType.Text = "";
            lblAutoID.Text = "";
        }

        public void clearAllVendorPart()
        {
            txtLeadTimeAir.Text = "";
            txtLeadTimeOcean.Text = "";
            txtLeadTimeGround.Text = "";
            txtPartDesc.Text = "";
            txtPartNum.Text = "";
            txtCurrLineAT.Text = "0";
            txtCurrLineNL.Text = "0";
            txtCurrLineSG.Text = "0";
            txtCurrPOSG.Text = "0";
            txtCurrPOAT.Text = "0";
            txtCurrPONL.Text = "0";
            // txtFactFob.Text = "";
            txtMIMOrdQty.Text = "0";
            txtOrdMultiPle.Text = "0";
            txtStdBoxQTY.Text = "0";
            txtStdPallQTY.Text = "0";
            this.cmbPlanType.SelectedIndex = 0;
            this.cmbShipMethod.SelectedIndex = 0;
            this.cmbFOB.SelectedIndex = 0;
            chkActive.Checked = true;
            lblPartVendID.Text = "";

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void grdVendor_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button5_Click_1(object sender, EventArgs e)
        {


        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            clearAllVendor();
        }

        private void grdVendPart_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex == 16)
            {

                DataGridViewTextBoxCell buttonEditCell = grdVendPart.Rows[e.RowIndex].Cells["btnEdit"] as DataGridViewTextBoxCell;
                buttonEditCell.Value = "EDIT";
                buttonEditCell.Style.ForeColor = Color.Blue;

            }
            //if (e.RowIndex > -1 && e.ColumnIndex == 6)
            //{
            //    DataGridViewCheckBoxCell inactiveCell = grdVendPart.Rows[e.RowIndex].Cells["chkDelete"] as DataGridViewCheckBoxCell;
            //   // DataGridViewTextBoxCell buttonDeleteCell = grdVendPart.Rows[e.RowIndex].Cells["btnDelete"] as DataGridViewTextBoxCell;
            //    //buttonDeleteCell.Value = "INACTIVE";
            //    inactiveCell.
            //    if (grdVendPart.Rows[e.RowIndex].Cells["active"].Value.ToString() == "Y")

            //        inactiveCell.FalseValue="y
            //    else
            //         inactiveCell.Selected=false;
            //    //buttonDeleteCell.Style.ForeColor = Color.Red;
            //   // if grdVendPart.Rows[e.RowIndex][e.ColumnIndex]


            //}
        }

        private void grdVendPart_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.ColumnIndex == 8)
            //{
            //    int auto_id = Convert.ToInt32(this.grdVendPart.Rows[e.RowIndex].Cells[0].Value);
            //    deleteVendPartRow(auto_id);
            //}
            if (e.ColumnIndex == 21)
            {
                lblPartVendID.Text = this.grdVendPart.Rows[e.RowIndex].Cells[0].Value.ToString();
                txtPartNum.Text = this.grdVendPart.Rows[e.RowIndex].Cells[2].Value.ToString();
                txtPartDesc.Text = this.grdVendPart.Rows[e.RowIndex].Cells[3].Value.ToString();


                if (this.grdVendPart.Rows[e.RowIndex].Cells[4].Value.ToString() == "AIR")
                    this.cmbShipMethod.SelectedIndex = 1;
                else if (this.grdVendPart.Rows[e.RowIndex].Cells[4].Value.ToString() == "OCEAN")
                    this.cmbShipMethod.SelectedIndex = 3;
                else if (this.grdVendPart.Rows[e.RowIndex].Cells[4].Value.ToString() == "GROUND")
                    this.cmbShipMethod.SelectedIndex = 2;
                txtLeadTimeAir.Text = this.grdVendPart.Rows[e.RowIndex].Cells[5].Value.ToString();
                this.txtLeadTimeOcean.Text = this.grdVendPart.Rows[e.RowIndex].Cells[6].Value.ToString();
                this.txtLeadTimeGround.Text = this.grdVendPart.Rows[e.RowIndex].Cells[7].Value.ToString();
                if (this.grdVendPart.Rows[e.RowIndex].Cells[8].Value.ToString() == "G")
                    cmbPlanType.SelectedIndex = 1;
                if (this.grdVendPart.Rows[e.RowIndex].Cells[8].Value.ToString() == "M")
                    cmbPlanType.SelectedIndex = 2;

                txtCurrPOAT.Text = this.grdVendPart.Rows[e.RowIndex].Cells[9].Value.ToString();
                txtCurrLineAT.Text = this.grdVendPart.Rows[e.RowIndex].Cells[10].Value.ToString();
                txtCurrPONL.Text = this.grdVendPart.Rows[e.RowIndex].Cells[11].Value.ToString();
                txtCurrLineNL.Text = this.grdVendPart.Rows[e.RowIndex].Cells[12].Value.ToString();
                txtCurrPOSG.Text = this.grdVendPart.Rows[e.RowIndex].Cells[13].Value.ToString();
                txtCurrLineSG.Text = this.grdVendPart.Rows[e.RowIndex].Cells[14].Value.ToString();
                this.txtStdBoxQTY.Text = this.grdVendPart.Rows[e.RowIndex].Cells[15].Value.ToString();
                this.txtStdPallQTY.Text = this.grdVendPart.Rows[e.RowIndex].Cells[16].Value.ToString();
                this.txtOrdMultiPle.Text = this.grdVendPart.Rows[e.RowIndex].Cells[17].Value.ToString();
                this.txtMIMOrdQty.Text = this.grdVendPart.Rows[e.RowIndex].Cells[18].Value.ToString();
                if (this.grdVendPart.Rows[e.RowIndex].Cells[19].Value.ToString() == "ASIA")

                    this.cmbFOB.SelectedIndex = 1;
                else if (this.grdVendPart.Rows[e.RowIndex].Cells[19].Value.ToString() == "EUROPE")

                    this.cmbFOB.SelectedIndex = 2;

                else if (this.grdVendPart.Rows[e.RowIndex].Cells[19].Value.ToString() == "US")

                    this.cmbFOB.SelectedIndex = 3;


                //this.txtFactFob.Text = this.grdVendPart.Rows[e.RowIndex].Cells[16].Value.ToString();


                if (this.grdVendPart.Rows[e.RowIndex].Cells[20].Value.ToString() == "ACTIVE")
                    chkActive.Checked = true;
                else
                    chkActive.Checked = false;
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {



        }

        private void btnLoadVendorPart_Click(object sender, EventArgs e)
        {

            this.loadShipMethod();
            loadPlanType();
            this.loadVendorPartMaster();
            this.loadFOB();
            clearAllVendorPart();
        }

        private void cmbVendor_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnClearVendPart_Click(object sender, EventArgs e)
        {
            clearAllVendorPart();
        }

        private void btnUpdateVendorPart_Click(object sender, EventArgs e)
        {
            string sqlStr = "";
            string act = "N";
            if (chkActive.Checked == true)
                act = "Y";
            if (cmbShipMethod.SelectedIndex == 0)
                MessageBox.Show("Please select the ship method");
            else if (cmbFOB.SelectedIndex == 0)
                MessageBox.Show("Please select the FOB");
            else
            {
                if (this.lblPartVendID.Text == "" || this.lblPartVendID.Text == "0")
                {
                    sqlStr = "insert into vc_vend_part(vend_id, part_num, short_desc, std_ship_meth,lead_time_air,lead_time_ocean,lead_time_ground,plan_type, curr_po_at, curr_line_at , curr_po_nl, curr_line_nl, curr_po_sg, curr_line_sg, std_box_qty, std_pallet_qty, ord_multiple, min_ord_qty, factory_fob) values('" + this.cmbVendPart.SelectedValue + "','" + this.txtPartNum.Text + "','" + this.txtPartDesc.Text + "','" + this.cmbShipMethod.Text + "','" + this.txtLeadTimeAir.Text + "','" + this.txtLeadTimeOcean.Text + "','" + this.txtLeadTimeGround.Text + "','" + this.cmbPlanType.SelectedValue.ToString() + "','" + this.txtCurrPOAT.Text + "','" + this.txtCurrLineAT.Text + "','" + this.txtCurrPONL.Text + "','" + this.txtCurrLineNL.Text + "','" + this.txtCurrPOSG.Text + "','" + this.txtCurrLineSG.Text + "','" + txtStdBoxQTY.Text + "','" + txtStdPallQTY.Text + "','" + txtOrdMultiPle.Text + "','" + txtMIMOrdQty.Text + "','" + cmbFOB.SelectedValue + "')";

                }
                else
                {
                    sqlStr = "update vc_vend_part set part_num='" + txtPartNum.Text + "',short_desc='" + txtPartDesc.Text + "', std_ship_meth='" + cmbShipMethod.Text + "',lead_time_air='" + txtLeadTimeAir.Text + "', lead_time_ocean='" + txtLeadTimeOcean.Text + "', lead_time_ground='" + this.txtLeadTimeGround.Text + "',curr_po_at='" + txtCurrPOAT.Text + "', curr_line_at='" + txtCurrLineAT.Text +"',curr_po_sg='" + txtCurrPOSG.Text + "', curr_line_sg='" + txtCurrLineSG.Text + "', curr_po_nl='" + txtCurrPONL.Text + "',plan_type='" + this.cmbPlanType.SelectedValue.ToString() + "', curr_line_nl='" + txtCurrLineNL.Text + "',  std_box_qty='" + txtStdBoxQTY.Text + "',std_pallet_qty='" + txtStdPallQTY.Text + "',ord_multiple='" + txtOrdMultiPle.Text + "',  min_ord_qty='" + txtMIMOrdQty.Text + "',  factory_fob ='" + cmbFOB.SelectedValue + "',Active='" + act + "'  where     vend_part_id=" + lblPartVendID.Text.Trim();
                }

                try
                {
                    SqlConnection cn = new SqlConnection(constr);
                    cn.Open();
                    SqlCommand cmd = new SqlCommand(sqlStr, cn);
                    cmd.ExecuteNonQuery();
                    cn.Close();
                    MessageBox.Show("Vendor has  been updated Successfully");
                    clearAllVendorPart();
                    loadVendorPartMaster();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Part has not been inserted due to some error. Please check"+ ex.Message.ToString());

                }
            }
        }

        private void btnUpdateVendor_Click(object sender, EventArgs e)
        {
            string sqlStr = "";
            sqlStr = " if exists(select 1 from vc_vend_master where ep_vend_code='" + txtVendorCode.Text.Trim() + "'  ) update vc_vend_master set ep_vend_code='" + txtVendorCode.Text + "',shared_doc_name='" + txtVendorType.Text + "' where ep_vend_code='" + txtVendorCode.Text.Trim() + "' else insert into vc_vend_master(ep_vend_code,shared_doc_name) values('" + txtVendorCode.Text + "','" + txtVendorType.Text + "') ";
         

            try
            {
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlCommand cmd = new SqlCommand(sqlStr, cn);
                cmd.ExecuteNonQuery();
                cn.Close();
                MessageBox.Show("Vendor has  been updated Successfully");
                clearAllVendor();
                loadVendorMaster();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Vendor has not been deleted due to some error. Please check");

            }
        }

        private void btnClearVendor_Click(object sender, EventArgs e)
        {
            clearAllVendor();
            //MessageBox.Show(textBox1.Text);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {


        }
        public void loadShipMethod()
        {
            string strsqlShipMthod = " select distinct std_ship_meth as disp,std_ship_meth as value from dbo.vc_vend_part where active='Y'";

            DataSet dsVendorShipMethod = getdataSet(strsqlShipMthod);
            DataRow dr = dsVendorShipMethod.Tables[0].NewRow();
            dr["value"] = "0";
            dr["disp"] = "Select";
            dsVendorShipMethod.Tables[0].Rows.InsertAt(dr, 0);
            cmbShipMethod.DataSource = dsVendorShipMethod.Tables[0];
            cmbShipMethod.DisplayMember = "disp";
            cmbShipMethod.ValueMember = "value";
        }

        public void loadPlanType()
        {
            DataSet dsVendorShipMethod = new DataSet();
            DataTable dt = new DataTable();
            dt.Columns.Add("Disp");
            dt.Columns.Add("value");

            DataRow dr0 = dt.NewRow();
            dr0["Disp"] = "Select";
            dr0["value"] = "0";
            dt.Rows.Add(dr0);

            DataRow dr = dt.NewRow();
            dr["Disp"] = "GOOGLE";
            dr["value"] = "G";
            dt.Rows.Add(dr);

            DataRow dr1 = dt.NewRow();
            dr1["Disp"] = "MIM";
            dr1["value"] = "M";
            dt.Rows.Add(dr1);

            // dsVendorShipMethod.Tables[0].Rows.InsertAt(dr, 0);
            this.cmbPlanType.DataSource = dt;
            cmbPlanType.DisplayMember = "disp";
            cmbPlanType.ValueMember = "value";

        }

        public void loadFOB()
        {

            DataSet dsVendorShipMethod = new DataSet();
            DataTable dt = new DataTable();
            dt.Columns.Add("Disp");
            dt.Columns.Add("value");

            DataRow dr0 = dt.NewRow();
            dr0["Disp"] = "Select";
            dr0["value"] = "0";
            dt.Rows.Add(dr0);

            DataRow dr = dt.NewRow();
            dr["Disp"] = "ASIA";
            dr["value"] = "ASIA";
            dt.Rows.Add(dr);

            DataRow dr1 = dt.NewRow();
            dr1["Disp"] = "EUROPE";
            dr1["value"] = "EUROPE";
            dt.Rows.Add(dr1);

            DataRow dr2 = dt.NewRow();
            dr2["Disp"] = "US";
            dr2["value"] = "US";
            dt.Rows.Add(dr2);

            // dsVendorShipMethod.Tables[0].Rows.InsertAt(dr, 0);
            this.cmbFOB.DataSource = dt;
            cmbFOB.DisplayMember = "disp";
            cmbFOB.ValueMember = "value";

        }

        public string Puntos(string strValor, int intNumDecimales)
        {

            string strAux = null;
            string strComas = null;
            string strPuntos = null;
            int intX = 0;
            bool bolMenos = false;

            strComas = "";
            if (strValor.Length == 0) return "";
            strValor = strValor.Replace(Application.CurrentCulture.NumberFormat.NumberGroupSeparator, "");
            if (strValor.Contains(Application.CurrentCulture.NumberFormat.NumberDecimalSeparator))
            {
                strAux = strValor.Substring(0, strValor.LastIndexOf(Application.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                strComas = strValor.Substring(strValor.LastIndexOf(Application.CurrentCulture.NumberFormat.NumberDecimalSeparator) + 1);
            }
            else
            {
                strAux = strValor;
            }

            if (strAux.Substring(0, 1) == Application.CurrentCulture.NumberFormat.NegativeSign)
            {
                bolMenos = true;
                strAux = strAux.Substring(1);
            }

            strPuntos = strAux;
            strAux = "";
            while (strPuntos.Length > 3)
            {
                strAux = Application.CurrentCulture.NumberFormat.NumberGroupSeparator + strPuntos.Substring(strPuntos.Length - 3, 3) + strAux;
                strPuntos = strPuntos.Substring(0, strPuntos.Length - 3);
            }
            if (intNumDecimales > 0)
            {
                if (strValor.Contains(Application.CurrentCulture.NumberFormat.PercentDecimalSeparator))
                {
                    strComas = Application.CurrentCulture.NumberFormat.PercentDecimalSeparator + strValor.Substring(strValor.LastIndexOf(Application.CurrentCulture.NumberFormat.PercentDecimalSeparator) + 1);
                    if (strComas.Length > intNumDecimales)
                    {
                        strComas = strComas.Substring(0, intNumDecimales + 1);
                    }

                }
            }
            strAux = strPuntos + strAux + strComas;


            return strAux;
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            //textBox1.Text = Puntos(textBox1.Text, 2);
            //textBox1.Select(textBox1.TextLength, 0);
        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        void tb_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar))
                e.Handled = true;

        }

        private void grdVendPart_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void cmbVendPart_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void txtStdPallQTY_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click_2(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }
        string filepath;
        //private void button5_Click_2(object sender, EventArgs e)
        //{
        //    Cursor.Current = Cursors.WaitCursor;
        //    if (lblstatusAT.Text == "READY TO SEND" && lblstatusNL.Text == "READY TO SEND")
        //    {

         
        //        ExcelObjectClass cls = new ExcelObjectClass();
        //        int wk = 0;
        //        if (lblWeek.Text != "")
        //        {
        //            wk = Convert.ToInt32(lblWeek.Text.ToString());
        //        }

        //        cls.cellUpdate(this.cmbVendorFileCreate.SelectedValue.ToString(), wk);

        //        filepath = cls.filename;
        //    }
        //    else
        //    {
        //        string message = "Vendor sheet for AT or NL hasn't been created for this week. Do you want to create shared doc?";
        //        string caption = "Create Shared Doc";
        //        MessageBoxButtons buttons = MessageBoxButtons.YesNo;
        //        DialogResult result;
        //        result = MessageBox.Show(message, caption, buttons);
        //        if (result == DialogResult.Yes)
        //        {
        //            ExcelObjectClass cls = new ExcelObjectClass();
        //            int wk = 0;
        //            if (lblWeek.Text != "")
        //            {
        //                wk = Convert.ToInt32(lblWeek.Text.ToString());
        //            }

        //            cls.cellUpdate(this.cmbVendorFileCreate.SelectedValue.ToString(), wk);

        //            filepath = cls.filename;

        //        }


        //    }
        //    Cursor.Current = Cursors.Default;
        //}

        //private void cmbVendorFileCreate_SelectedIndexChanged(object sender, EventArgs e)
        //{

        //    clearFileCreateAttribute();
        //    if (cmbVendorFileCreate.SelectedIndex > 0)
        //    {
        //        string sqlqry = "select distinct site,convert(varchar,ready_to_send_date,101) as create_date,dbo.udf_GetISOWeekNumberFromDate(ready_to_send_date) as create_week ,case when status='P' then 'READY TO SEND' else '' end  as status  from vc_commit where vend_id=" + cmbVendorFileCreate.SelectedValue + " and status='P'";
        //        DataSet ds = getdataSet(sqlqry);
        //        if (ds != null)
        //            if (ds.Tables.Count > 0)
        //                if (ds.Tables[0].Rows.Count > 0)
        //                    lblWeek.Text = ds.Tables[0].Rows[0]["create_week"].ToString();
        //        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
        //        {

        //            if (ds.Tables[0].Rows[i]["site"].ToString() == "AT")
        //            {
        //                this.updateDateAT.Text = ds.Tables[0].Rows[i]["create_date"].ToString();
        //                this.lblstatusAT.Text = ds.Tables[0].Rows[i]["status"].ToString();
        //            }

        //            else if (ds.Tables[0].Rows[i]["site"].ToString() == "NL")
        //            {
        //                this.updateDateNL.Text = ds.Tables[0].Rows[i]["create_date"].ToString();
        //                this.lblstatusNL.Text = ds.Tables[0].Rows[i]["status"].ToString();
        //            }
        //        }
        //    }
        //}



        //public string getCollectionName()
        //{
        //    string coll_name = "";
        //    string sql = "select * from dbo.vc_vend_master where vend_id=" + cmbVendorFileCreate.SelectedValue;
        //    DataSet ds = getdataSet(sql);
        //    if (ds != null)
        //        if (ds.Tables.Count > 0)
        //            if (ds.Tables[0].Rows.Count > 0)
        //                coll_name = ds.Tables[0].Rows[0]["shared_doc_name"].ToString();


        //    return coll_name;
        //}

      //  public static Google.Apis.Drive.v2.DriveService dr { get; private set; }
      //  static String CLIENT_ID = "943908261643.apps.googleusercontent.com";
      //  static String CLIENT_SECRET = "63oqKhgS7Gm9I3mx3PBVGcXn";
      //  static String REDIRECT_URI = "rn:ietf:wg:oauth:2.0:oob http://localhost";
      //  static String[] SCOPES = new String[] {  "https://www.googleapis.com/auth/drive.file",
      //"https://www.googleapis.com/auth/userinfo.email",
      //"https://www.googleapis.com/auth/userinfo.profile"  
      //  };
        //private void btnUpdateSharedDoc_Click(object sender, EventArgs e)
        //{
        //    Cursor.Current = Cursors.WaitCursor;

        //    dr = new Google.Apis.Drive.v2.DriveService(CreateAuthenticator());
        //    if (dr != null)
        //    {
        //        List<Google.Apis.Drive.v2.Data.File> fileList = Google.Apis.Util.Utilities.RetrieveAllFiles(dr);

        //        OpenFileDialog dialog = new OpenFileDialog();
        //        //dialog.AddExtension = true;
        //        //dialog.DefaultExt = ".txt";
        //        //dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
        //        //dialog.Multiselect = false;
        //        //dialog.ShowDialog();
        //        //dialog.OpenFile(.OpenFile(
        //        string Parentid = "";
        //        foreach (Google.Apis.Drive.v2.Data.File item in fileList)
        //        {
        //            if (item.Title == getCollectionName())
        //            {
        //                Parentid = item.Id;
        //                break;
        //            }
        //        }
        //        File body = new File();
        //        body.Title = System.IO.Path.GetFileName(filepath);//dialog.FileName);
        //        body.Description = "document";
        //        body.MimeType = "application/xlsx";


        //        if (!String.IsNullOrEmpty(Parentid))
        //        {
        //            body.Parents = new List<ParentReference>() { new ParentReference() { Id = Parentid } };
        //        }

        //        System.IO.Stream fileStream = System.IO.File.Open(filepath.Replace("file:\\", ""), System.IO.FileMode.Open, System.IO.FileAccess.Read);
        //        byte[] byteArray = new byte[fileStream.Length];
        //        fileStream.Read(byteArray, 0, (int)fileStream.Length);

        //        System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
        //        FilesResource.InsertMediaUpload request = dr.Files.Insert(body, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
        //        request.Convert = true;

        //        request.Upload();
        //        MessageBox.Show("File has been uploaded in shared doc");
        //        Cursor.Current = Cursors.Default;
        //    }
        //}

      



        private void panel14_Paint(object sender, PaintEventArgs e)
        {

        }

        //private void btnVCLoadGrid_Click(object sender, EventArgs e)
        //{
        //    Cursor.Current = Cursors.WaitCursor;
        //    if (cmbVCVendor.SelectedIndex == 0)
        //    {
        //        MessageBox.Show("Please select vendor ");
        //    }
        //    else
        //    {              
        //        getVendorPart(cmbVCVendor.SelectedValue.ToString());
        //    }
        //    Cursor.Current = Cursors.Default;
        //}



        DataSet dsVendorCommitPart;
        DataTable distinctValues_vc;
        //public void getVendorPart(string vendor)
        //{

        //     string strsqlPart = " select * from m_vc_requested_qty where vend_id= " + cmbVCVendor.SelectedValue ;/// select part_num,short_desc from vc_vend_part where vend_id=" + cmbVendor.SelectedValue;
        //    //try
        //    //  {
        //        dsVendorCommitPart = getdataSet(strsqlPart);
        //        DataSet dsWeek = getdataSet("select top 20 convert(varchar,date,101) as date,mweek,myear from m_week_number_table where date >=(getdate()-7) ");
        //    // DataSet dsCommitReq = getdataSet("select * from vc_commit where status='P' and vend_id='" + cmbVendor.SelectedValue + "' and date>getdate()-42 order by year,week");
        //    //   DataSet dsOffset = getdataSet("select *,dbo.udf_GetISOWeekNumberFromDate(dateadd(day, vend.lead_time_air,wk.date)) air_time_week,dbo.udf_GetISOWeekNumberFromDate(dateadd(day, vend.lead_time_ocean,wk.date)) as ocean_time_week from (select part_num,lead_time_air,lead_time_ocean from vc_vend_part   where vend_id=" + this.cmbVendor.SelectedValue + " ) as vend cross join (select top 26  convert(varchar,date,101) as date,mweek,myear from m_week_number_table where date >=(getdate()-49)) as wk");
        //    //   lblCommitId.Text = "";
        //    customcontrol.VendorCommitHeader ctrlHeader = new VendorCommitHeader();
        //    ctrlHeader.Name = "header";
        //    ctrlHeader.Width = 1600;
            
           
        //    ctrlHeader.lblCurrentWk.Text = dsWeek.Tables[0].Rows[0]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear.Text = dsWeek.Tables[0].Rows[0]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate.Text = dsWeek.Tables[0].Rows[0]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus1.Text = dsWeek.Tables[0].Rows[1]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_1.Text = dsWeek.Tables[0].Rows[1]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_1.Text = dsWeek.Tables[0].Rows[1]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus2.Text = dsWeek.Tables[0].Rows[2]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_2.Text = dsWeek.Tables[0].Rows[2]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_2.Text = dsWeek.Tables[0].Rows[2]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus3.Text = dsWeek.Tables[0].Rows[3]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_3.Text = dsWeek.Tables[0].Rows[3]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_3.Text = dsWeek.Tables[0].Rows[3]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus4.Text = dsWeek.Tables[0].Rows[4]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_4.Text = dsWeek.Tables[0].Rows[4]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_4.Text = dsWeek.Tables[0].Rows[4]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus5.Text = dsWeek.Tables[0].Rows[5]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_5.Text = dsWeek.Tables[0].Rows[5]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_5.Text = dsWeek.Tables[0].Rows[5]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus6.Text = dsWeek.Tables[0].Rows[6]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_6.Text = dsWeek.Tables[0].Rows[6]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_6.Text = dsWeek.Tables[0].Rows[6]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus7.Text = dsWeek.Tables[0].Rows[7]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_7.Text = dsWeek.Tables[0].Rows[7]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_7.Text = dsWeek.Tables[0].Rows[7]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus8.Text = dsWeek.Tables[0].Rows[8]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_8.Text = dsWeek.Tables[0].Rows[8]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_8.Text = dsWeek.Tables[0].Rows[8]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus9.Text = dsWeek.Tables[0].Rows[9]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_9.Text = dsWeek.Tables[0].Rows[9]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_9.Text = dsWeek.Tables[0].Rows[9]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus10.Text = dsWeek.Tables[0].Rows[10]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_10.Text = dsWeek.Tables[0].Rows[10]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_10.Text = dsWeek.Tables[0].Rows[10]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus11.Text = dsWeek.Tables[0].Rows[11]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_11.Text = dsWeek.Tables[0].Rows[11]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_11.Text = dsWeek.Tables[0].Rows[11]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus12.Text = dsWeek.Tables[0].Rows[12]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_12.Text = dsWeek.Tables[0].Rows[12]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_12.Text = dsWeek.Tables[0].Rows[12]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus13.Text = dsWeek.Tables[0].Rows[13]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_13.Text = dsWeek.Tables[0].Rows[13]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_13.Text = dsWeek.Tables[0].Rows[13]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus14.Text = dsWeek.Tables[0].Rows[14]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_14.Text = dsWeek.Tables[0].Rows[14]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_14.Text = dsWeek.Tables[0].Rows[14]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus15.Text = dsWeek.Tables[0].Rows[15]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_15.Text = dsWeek.Tables[0].Rows[15]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_15.Text = dsWeek.Tables[0].Rows[15]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus16.Text = dsWeek.Tables[0].Rows[16]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_16.Text = dsWeek.Tables[0].Rows[16]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_16.Text = dsWeek.Tables[0].Rows[16]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus17.Text = dsWeek.Tables[0].Rows[17]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_17.Text = dsWeek.Tables[0].Rows[17]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_17.Text = dsWeek.Tables[0].Rows[17]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus18.Text = dsWeek.Tables[0].Rows[18]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_18.Text = dsWeek.Tables[0].Rows[18]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_18.Text = dsWeek.Tables[0].Rows[18]["date"].ToString();

        //    ctrlHeader.lblCurrentWk_plus19.Text = dsWeek.Tables[0].Rows[19]["mweek"].ToString();
        //    ctrlHeader.lblCurrentYear_plus_19.Text = dsWeek.Tables[0].Rows[19]["myear"].ToString();
        //    ctrlHeader.lblCurrentWkDate_plus_19.Text = dsWeek.Tables[0].Rows[19]["date"].ToString();

        //    TableLayoutPanel tableLayoutPanelheader = new TableLayoutPanel();

        //    tableLayoutPanelheader.Width = pnlHeader.Width;
        //    tableLayoutPanelheader.Height = pnlHeader.Height;

        //    tableLayoutPanelheader.Font = new Font("Verdana", 8, FontStyle.Regular);
        //    ////tableLayoutPanel1.CellBorderStyle = CellBorderStyle.FixedSingle;
        //    tableLayoutPanelheader.AutoScroll = true;
        //    tableLayoutPanelheader.ColumnCount = 2;
        //    for (int i = 0; i < tableLayoutPanelheader.ColumnCount; i++)
        //    {
        //        if (i == 0)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 200);
        //            tableLayoutPanelheader.ColumnStyles.Add(cs);
        //        }
             
        //        if (i == 1)
        //        {
        //            ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 1600);
        //            tableLayoutPanelheader.ColumnStyles.Add(cs);
        //        }
        //                 }
        //    //if (dsPart != null)
        //    //{
        //    //    //dataGridView1.DataSource = dsPart.Tables[0];
        //    //    for (int j = 0; j < panel2.Controls.Count; j++)
        //    //    {
        //    //        panel14.Controls.RemoveAt(j);
        //    //    }
        //    //}

        //        ////tableLayoutPanelheader.RowCount = dsPart.Tables[0].Rows.Count+1;
        //        Label lblPart = new Label();
        //        lblPart.Text = "PART NUM";
        //        lblPart.Width = 200;
        //        lblPart.TextAlign = ContentAlignment.MiddleCenter;

        //       tableLayoutPanelheader.Controls.Add(lblPart, 0, 0);
               
        //        tableLayoutPanelheader.Controls.Add(ctrlHeader, 1, 0);
        //        pnlVCHeader.Controls.Add(tableLayoutPanelheader);
        //        TableLayoutPanel tableLayoutPanel1 = new TableLayoutPanel();
        //        if (dsVendorCommitPart != null)
        //        {
        //            //dataGridView1.DataSource = dsPart.Tables[0];
        //            for (int j = 0; j < panel14.Controls.Count; j++)
        //            {
        //                panel14.Controls.RemoveAt(j);
        //            }

        //                    DataView view_vc = new DataView(dsVendorCommitPart.Tables[0]);
        //                    distinctValues_vc = view_vc.ToTable(true, "part_num");
        //                    int partcount_vc = distinctValues_vc.Rows.Count;
        //                    String[] leadtime;
        //        for (int rownum = 0; rownum < partcount_vc; rownum++)
        //        {
        //            DataRow[] drPartATRequest = dsVendorCommitPart.Tables[0].Select("part_num='" + distinctValues_vc.Rows[rownum]["part_num"].ToString() + "' and site='AT'");
        //            DataRow[] drPartNLRequest = dsVendorCommitPart.Tables[0].Select("part_num='" + distinctValues_vc.Rows[rownum]["part_num"].ToString() + "' and site='NL'");            
                     
                       
        //                tableLayoutPanel1.Width = panel2.Width;
        //                tableLayoutPanel1.Height = panel2.Height;

        //                tableLayoutPanel1.Font = new Font("Verdana", 8, FontStyle.Regular);
        //                //tableLayoutPanel1.CellBorderStyle = CellBorderStyle.FixedSingle;
        //                tableLayoutPanel1.AutoScroll = true;
        //                tableLayoutPanel1.ColumnCount = 2;
        //                for (int i = 0; i < tableLayoutPanel1.ColumnCount; i++)
        //                {
        //                    if (i == 0)
        //                    {
        //                        ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 200);
        //                        tableLayoutPanel1.ColumnStyles.Add(cs);
        //                    }

        //                    if (i == 1)
        //                    {
        //                        ColumnStyle cs = new ColumnStyle(SizeType.Absolute, 1600);
        //                        tableLayoutPanel1.ColumnStyles.Add(cs);
        //                    }
        //                }
                                      
        //                customcontrol.ctrlVendorCommit ctrl = new customcontrol.ctrlVendorCommit();
        //                ctrl.Name = "req" + distinctValues_vc.Rows[rownum]["part_num"].ToString();
        //                ctrl.Width = 1600;
        //                Label lblPartName = new Label();
        //                lblPartName.Width = 200;
        //                lblPartName.TextAlign = ContentAlignment.MiddleCenter;
        //                lblPartName.Text = distinctValues_vc.Rows[rownum]["part_num"].ToString();

        //                if (drPartATRequest.Length > 0)
        //                {
        //                    ctrl.lblCom_AT.Text = drPartATRequest[0]["com_id"].ToString();
        //                    ctrl.txtNewSupp_wk1.Text = drPartATRequest[0]["new_req_qty_1"].ToString();
        //                    ctrl.txtNewSupp_wk2.Text = drPartATRequest[0]["new_req_qty_2"].ToString();
        //                    ctrl.txtNewSupp_wk3.Text = drPartATRequest[0]["new_req_qty_3"].ToString();
        //                    ctrl.txtNewSupp_wk4.Text = drPartATRequest[0]["new_req_qty_4"].ToString();
        //                    ctrl.txtNewSupp_wk5.Text = drPartATRequest[0]["new_req_qty_5"].ToString();
        //                    ctrl.txtNewSupp_wk6.Text = drPartATRequest[0]["new_req_qty_6"].ToString();
        //                    ctrl.txtNewSupp_wk7.Text = drPartATRequest[0]["new_req_qty_7"].ToString();
        //                    ctrl.txtNewSupp_wk8.Text = drPartATRequest[0]["new_req_qty_8"].ToString();
        //                    ctrl.txtNewSupp_wk9.Text = drPartATRequest[0]["new_req_qty_9"].ToString();
        //                    ctrl.txtNewSupp_wk10.Text = drPartATRequest[0]["new_req_qty_10"].ToString();
        //                    ctrl.txtNewSupp_wk11.Text = drPartATRequest[0]["new_req_qty_11"].ToString();
        //                    ctrl.txtNewSupp_wk12.Text = drPartATRequest[0]["new_req_qty_12"].ToString();
        //                    ctrl.txtNewSupp_wk13.Text = drPartATRequest[0]["new_req_qty_13"].ToString();
        //                    ctrl.txtNewSupp_wk14.Text = drPartATRequest[0]["new_req_qty_14"].ToString();
        //                    ctrl.txtNewSupp_wk15.Text = drPartATRequest[0]["new_req_qty_15"].ToString();
        //                    ctrl.txtNewSupp_wk16.Text = drPartATRequest[0]["new_req_qty_16"].ToString();
        //                    ctrl.txtNewSupp_wk17.Text = drPartATRequest[0]["new_req_qty_17"].ToString();
        //                    ctrl.txtNewSupp_wk18.Text = drPartATRequest[0]["new_req_qty_18"].ToString();
        //                    ctrl.txtNewSupp_wk19.Text = drPartATRequest[0]["new_req_qty_19"].ToString();
        //                    ctrl.txtNewSupp_wk20.Text = drPartATRequest[0]["new_req_qty_20"].ToString();
        //                }

        //                if (drPartNLRequest.Length > 0)
        //                {
        //                    ctrl.lblcom_NL.Text = drPartNLRequest[0]["com_id"].ToString();
        //                    ctrl.txtNewSupp_wk1_NL.Text = drPartNLRequest[0]["new_req_qty_1"].ToString();
        //                    ctrl.txtNewSupp_wk2_NL.Text = drPartNLRequest[0]["new_req_qty_2"].ToString();
        //                    ctrl.txtNewSupp_wk3_NL.Text = drPartNLRequest[0]["new_req_qty_3"].ToString();
        //                    ctrl.txtNewSupp_wk4_NL.Text = drPartNLRequest[0]["new_req_qty_4"].ToString();
        //                    ctrl.txtNewSupp_wk5_NL.Text = drPartNLRequest[0]["new_req_qty_5"].ToString();
        //                    ctrl.txtNewSupp_wk6_NL.Text = drPartNLRequest[0]["new_req_qty_6"].ToString();
        //                    ctrl.txtNewSupp_wk7_NL.Text = drPartNLRequest[0]["new_req_qty_7"].ToString();
        //                    ctrl.txtNewSupp_wk8_NL.Text = drPartNLRequest[0]["new_req_qty_8"].ToString();
        //                    ctrl.txtNewSupp_wk9_NL.Text = drPartNLRequest[0]["new_req_qty_9"].ToString();
        //                    ctrl.txtNewSupp_wk10_NL.Text = drPartNLRequest[0]["new_req_qty_10"].ToString();
        //                    ctrl.txtNewSupp_wk11_NL.Text = drPartNLRequest[0]["new_req_qty_11"].ToString();
        //                    ctrl.txtNewSupp_wk12_NL.Text = drPartNLRequest[0]["new_req_qty_12"].ToString();
        //                    ctrl.txtNewSupp_wk13_NL.Text = drPartNLRequest[0]["new_req_qty_13"].ToString();
        //                    ctrl.txtNewSupp_wk14_NL.Text = drPartNLRequest[0]["new_req_qty_14"].ToString();
        //                    ctrl.txtNewSupp_wk15_NL.Text = drPartNLRequest[0]["new_req_qty_15"].ToString();
        //                    ctrl.txtNewSupp_wk16_NL.Text = drPartNLRequest[0]["new_req_qty_16"].ToString();
        //                    ctrl.txtNewSupp_wk17_NL.Text = drPartNLRequest[0]["new_req_qty_17"].ToString();
        //                    ctrl.txtNewSupp_wk18_NL.Text = drPartNLRequest[0]["new_req_qty_18"].ToString();
        //                    ctrl.txtNewSupp_wk19_NL.Text = drPartNLRequest[0]["new_req_qty_19"].ToString();
        //                    ctrl.txtNewSupp_wk20_NL.Text = drPartNLRequest[0]["new_req_qty_20"].ToString();
        //                }

        //                if (rownum % 2 == 0)
        //                {
        //                    lblPartName.BackColor = System.Drawing.Color.Bisque;                         
        //                    ctrl.BackColor = System.Drawing.Color.Bisque;
        //                }
        //                else
        //                {
        //                    lblPartName.BackColor = System.Drawing.Color.PeachPuff;
        //                    ctrl.BackColor = System.Drawing.Color.PeachPuff;
        //                }
              
        //                tableLayoutPanel1.Controls.Add(lblPartName, 0, rownum + 1);                                      
        //                tableLayoutPanel1.Controls.Add(ctrl, 1, rownum + 1);
        //            }
        //            tableLayoutPanel1.CellPaint += new TableLayoutCellPaintEventHandler(tableLayoutPanel1_CellPaint);
        //            panel14.Controls.Add(tableLayoutPanel1);
        //        }
            
        //    }

        //private void btnSaveVendorCommit_Click(object sender, EventArgs e)
        //{

        //    Cursor.Current = Cursors.WaitCursor;
        //    btnSaveVendorCommit.Enabled = false;
        //    string sql = "";
           
        //    string status = "";
        //    int com_id_AT = 0; //Convert.ToInt32(user.lblCom_AT.Text.ToString());
        //    int com_id_NL = 0; //Convert.ToInt32(user.lblcom_NL.Text.ToString());


        //    for (int r = 0; r < distinctValues_vc.Rows.Count; r++)
        //        {
        //            string part =   distinctValues_vc.Rows[r]["part_num"].ToString();
        //            // drDemand = dsPart.Tables[0].Select("part_num='" + distinctValues.Rows[r]["part_num"].ToString() + "'");
        //            //   DataRow[] drSupply = dsPart.Tables[1].Select("part_num='" + distinctValues.Rows[r]["part_num"].ToString() + "'");

        //            ctrlVendorCommit user = panel14.Controls.Find("req" + part, true)[0] as ctrlVendorCommit;
        //            VendorCommitHeader header = pnlVCHeader.Controls.Find("header", true)[0] as VendorCommitHeader;
                    

        //            if (user != null)
        //            {
        //                int vendor_qty_1_AT = 0;
        //                int vendor_qty_2_AT = 0;
        //                int vendor_qty_3_AT = 0;
        //                int vendor_qty_4_AT = 0;
        //                int vendor_qty_5_AT = 0;
        //                int vendor_qty_6_AT = 0;
        //                int vendor_qty_7_AT = 0;
        //                int vendor_qty_8_AT = 0;
        //                int vendor_qty_9_AT = 0;
        //                int vendor_qty_10_AT = 0;
        //                int vendor_qty_11_AT = 0;
        //                int vendor_qty_12_AT = 0;
        //                int vendor_qty_13_AT = 0;
        //                int vendor_qty_14_AT = 0;
        //                int vendor_qty_15_AT = 0;
        //                int vendor_qty_16_AT = 0;
        //                int vendor_qty_17_AT = 0;
        //                int vendor_qty_18_AT = 0;
        //                int vendor_qty_19_AT = 0;
        //                int vendor_qty_20_AT = 0;

        //                int vendor_qty_1_NL = 0;
        //                int vendor_qty_2_NL = 0;
        //                int vendor_qty_3_NL = 0;
        //                int vendor_qty_4_NL = 0;
        //                int vendor_qty_5_NL = 0;
        //                int vendor_qty_6_NL = 0;
        //                int vendor_qty_7_NL = 0;
        //                int vendor_qty_8_NL = 0;
        //                int vendor_qty_9_NL = 0;
        //                int vendor_qty_10_NL = 0;
        //                int vendor_qty_11_NL = 0;
        //                int vendor_qty_12_NL = 0;
        //                int vendor_qty_13_NL = 0;
        //                int vendor_qty_14_NL = 0;
        //                int vendor_qty_15_NL = 0;
        //                int vendor_qty_16_NL = 0;
        //                int vendor_qty_17_NL = 0;
        //                int vendor_qty_18_NL = 0;
        //                int vendor_qty_19_NL = 0;
        //                int vendor_qty_20_NL = 0;
                        
        //                if (user.txtATVendorQty1.Text != "")
        //                    vendor_qty_1_AT = Convert.ToInt32(user.txtATVendorQty1.Text);
                        
        //                if (user.txtATVendorQty2.Text != "")
        //                    vendor_qty_2_AT = Convert.ToInt32(user.txtATVendorQty2.Text);

        //                if (user.txtATVendorQty3.Text != "")
        //                    vendor_qty_3_AT = Convert.ToInt32(user.txtATVendorQty3.Text);

        //                if (user.txtATVendorQty4.Text != "")
        //                    vendor_qty_4_AT = Convert.ToInt32(user.txtATVendorQty4.Text);

        //                if (user.txtATVendorQty15.Text != "")
        //                    vendor_qty_5_AT = Convert.ToInt32(user.txtATVendorQty5.Text);

        //                if (user.txtATVendorQty16.Text != "")
        //                    vendor_qty_6_AT = Convert.ToInt32(user.txtATVendorQty6.Text);

        //                if (user.txtATVendorQty7.Text != "")
        //                    vendor_qty_7_AT = Convert.ToInt32(user.txtATVendorQty7.Text);

        //                if (user.txtATVendorQty8.Text != "")
        //                    vendor_qty_8_AT = Convert.ToInt32(user.txtATVendorQty8.Text);

        //                if (user.txtATVendorQty9.Text != "")
        //                    vendor_qty_9_AT = Convert.ToInt32(user.txtATVendorQty9.Text);

        //                if (user.txtATVendorQty10.Text != "")
        //                    vendor_qty_10_AT = Convert.ToInt32(user.txtATVendorQty10.Text);

        //                if (user.txtATVendorQty11.Text != "")
        //                    vendor_qty_11_AT = Convert.ToInt32(user.txtATVendorQty11.Text);

        //                if (user.txtATVendorQty12.Text != "")
        //                    vendor_qty_12_AT = Convert.ToInt32(user.txtATVendorQty12.Text);

        //                if (user.txtATVendorQty13.Text != "")
        //                    vendor_qty_13_AT = Convert.ToInt32(user.txtATVendorQty13.Text);

        //                if (user.txtATVendorQty14.Text != "")
        //                    vendor_qty_14_AT = Convert.ToInt32(user.txtATVendorQty14.Text);

        //                if (user.txtATVendorQty15.Text != "")
        //                    vendor_qty_15_AT = Convert.ToInt32(user.txtATVendorQty15.Text);

        //                if (user.txtATVendorQty16.Text != "")
        //                    vendor_qty_16_AT = Convert.ToInt32(user.txtATVendorQty16.Text);

        //                if (user.txtATVendorQty17.Text != "")
        //                    vendor_qty_17_AT = Convert.ToInt32(user.txtATVendorQty17.Text);

        //                if (user.txtATVendorQty18.Text != "")
        //                    vendor_qty_18_AT = Convert.ToInt32(user.txtATVendorQty18.Text);

        //                if (user.txtATVendorQty19.Text != "")
        //                    vendor_qty_19_AT = Convert.ToInt32(user.txtATVendorQty19.Text);

        //                if (user.txtATVendorQty20.Text != "")
        //                    vendor_qty_20_AT = Convert.ToInt32(user.txtATVendorQty20.Text);
                        
        //                if (user.txtNLVendorQty1.Text != "")
        //                    vendor_qty_1_NL = Convert.ToInt32(user.txtNLVendorQty1.Text);
                        
        //                if (user.txtNLVendorQty2.Text != "")
        //                    vendor_qty_2_NL = Convert.ToInt32(user.txtNLVendorQty2.Text);

        //                if (user.txtNLVendorQty3.Text != "")
        //                    vendor_qty_3_NL = Convert.ToInt32(user.txtNLVendorQty3.Text);

        //                if (user.txtNLVendorQty4.Text != "")
        //                    vendor_qty_4_NL = Convert.ToInt32(user.txtNLVendorQty4.Text);

        //                if (user.txtNLVendorQty15.Text != "")
        //                    vendor_qty_5_NL = Convert.ToInt32(user.txtNLVendorQty5.Text);

        //                if (user.txtNLVendorQty16.Text != "")
        //                    vendor_qty_6_NL = Convert.ToInt32(user.txtNLVendorQty6.Text);

        //                if (user.txtNLVendorQty7.Text != "")
        //                    vendor_qty_7_NL = Convert.ToInt32(user.txtNLVendorQty7.Text);

        //                if (user.txtNLVendorQty8.Text != "")
        //                    vendor_qty_8_NL = Convert.ToInt32(user.txtNLVendorQty8.Text);

        //                if (user.txtNLVendorQty9.Text != "")
        //                    vendor_qty_9_NL = Convert.ToInt32(user.txtNLVendorQty9.Text);

        //                if (user.txtNLVendorQty10.Text != "")
        //                    vendor_qty_10_NL = Convert.ToInt32(user.txtNLVendorQty10.Text);

        //                if (user.txtNLVendorQty11.Text != "")
        //                    vendor_qty_11_NL = Convert.ToInt32(user.txtNLVendorQty11.Text);

        //                if (user.txtNLVendorQty12.Text != "")
        //                    vendor_qty_12_NL = Convert.ToInt32(user.txtNLVendorQty12.Text);

        //                if (user.txtNLVendorQty13.Text != "")
        //                    vendor_qty_13_NL = Convert.ToInt32(user.txtNLVendorQty13.Text);


        //                if (user.txtNLVendorQty14.Text != "")
        //                    vendor_qty_14_NL = Convert.ToInt32(user.txtNLVendorQty14.Text);

        //                if (user.txtNLVendorQty15.Text != "")
        //                    vendor_qty_15_NL = Convert.ToInt32(user.txtNLVendorQty15.Text);

        //                if (user.txtNLVendorQty16.Text != "")
        //                    vendor_qty_16_NL = Convert.ToInt32(user.txtNLVendorQty16.Text);

        //                if (user.txtNLVendorQty17.Text != "")
        //                    vendor_qty_17_NL = Convert.ToInt32(user.txtNLVendorQty17.Text);

        //                if (user.txtNLVendorQty18.Text != "")
        //                    vendor_qty_18_NL = Convert.ToInt32(user.txtNLVendorQty18.Text);

        //                if (user.txtNLVendorQty19.Text != "")
        //                    vendor_qty_19_NL = Convert.ToInt32(user.txtNLVendorQty19.Text);

        //                if (user.txtNLVendorQty20.Text != "")
        //                    vendor_qty_20_NL = Convert.ToInt32(user.txtNLVendorQty20.Text);

                                                
                      
        //                if (user.lblCom_AT.Text.ToString()!="")
        //                {
        //                    com_id_AT = Convert.ToInt32(user.lblCom_AT.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_1_AT + "' , status='C',commit_date=getdate() where site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_2_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus1.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_3_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus2.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_4_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus3.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_5_AT + "' , status='C',commit_date=getdate()  where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus4.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_6_AT + "'  , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus5.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_7_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus6.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_8_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus7.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_9_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus8.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_10_AT + "', status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus9.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_11_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus10.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_12_AT + "', status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus11.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_13_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus12.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_14_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus13.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_15_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus14.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_16_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus15.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_17_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus16.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_18_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus17.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_19_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus18.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_20_AT + "' , status='C',commit_date=getdate() where  site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus19.Text.ToString());
        //                    sql = sql + "  update vc_commit set new_com_qty=prev_com_qty,ship_method=prev_ship_method,status='C',commit_date=getdate()  where   site='AT' and com_id=" + com_id_AT + " and part_num='" + part + "' and (status='P') and  date<'" + header.lblCurrentWkDate.Text.ToString() + "'";
                            
                                                    
        //                }

        //                if (user.lblcom_NL.Text.ToString()!="")
        //                {
        //                    com_id_NL = Convert.ToInt32(user.lblcom_NL.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_1_NL + "' , status='C',commit_date=getdate() where site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_2_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus1.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_3_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus2.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_4_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus3.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_5_NL + "' , status='C',commit_date=getdate()  where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus4.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_6_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus5.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_7_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus6.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_8_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus7.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_9_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus8.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_10_NL + "', status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus9.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_11_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus10.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_12_NL + "', status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus11.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_13_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus12.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_14_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus13.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_15_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus14.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_16_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus15.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_17_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus16.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_18_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus17.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_19_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus18.Text.ToString());
        //                    sql = sql + " update vc_commit set new_com_qty='" + vendor_qty_20_NL + "' , status='C',commit_date=getdate() where  site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and week=" + Convert.ToInt32(header.lblCurrentWk_plus19.Text.ToString());
        //                    sql = sql + "  update vc_commit set new_com_qty=prev_com_qty,ship_method=prev_ship_method,status='C',commit_date=getdate()  where   site='NL' and com_id=" + com_id_NL + " and part_num='" + part + "' and (status='P') and  date<'" + header.lblCurrentWkDate.Text.ToString() + "'";
        //                }
        //            }
        //        }
        //        try
        //        {
        //            SqlConnection cn = new SqlConnection(this.constr);
        //            cn.Open();
        //            sql = sql + "  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) select 'N',site,vend_id,getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),part_num,date,week,year,new_com_qty,ship_method,commit_name,plan_type,po,line from vc_commit where com_id=" + com_id_AT + " ";
        //            sql = sql + " update vc_commit set com_id =(select max(auto_id) from vc_commit) where com_id is null  insert into vc_commit(status,site,vend_id,create_date,create_week,part_num,date,week,year,prev_com_qty,prev_ship_method,commit_name,plan_type,po,line) select 'N',site,vend_id,getdate(), dbo.udf_GetISOWeekNumberFromDate(getdate()),part_num,date,week,year,new_com_qty,ship_method,commit_name,plan_type,po,line from vc_commit where com_id=" + com_id_NL + " ";
        //            sql = sql + "  update vc_commit set com_id =(select max(auto_id) from vc_commit) where com_id is null";
        //            SqlCommand cmd = new SqlCommand(sql, cn);
        //            cmd.CommandTimeout = 0;
        //            cmd.ExecuteNonQuery();
        //            button2.Enabled = false;
        //            button2.BackColor = System.Drawing.Color.IndianRed;
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Error :" + ex.Message.ToString());
        //        }
        //        MessageBox.Show(" Vendor Commit has been  commited  " );
        //        //getPart();
        //        btnSaveVendorCommit.Enabled = false;
        //        cmbVCVendor.SelectedIndex = 0;
        //        panel4.Controls.Clear();
        //        Cursor.Current = Cursors.Default;
        //    }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            //loadVendor();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
          //  loadVCVendor();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            loadVendorMaster();
        }
       
        }
    }


        
    





            
          
        
 

    


