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



//******************************************************************************************************************************************************************//

//Tables:vc_vend_part,vca_req_com,vc_vend_master
//views:vca_ship_tracker_for_vcapp(vca_distinct_ship_id,vca_ship_tracker,vca_tracking_links)
//Stored Proc:
//vca_ship_tracker_for_vcapp,vca_ship_tracker_not_received,vca_ship_tracker_received
//******************************************************************************************************************************************************************//

namespace Version3
{
    public partial class VendorCommitApp : Form
    {

        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        string constrTest = ConfigurationManager.ConnectionStrings["epicorTESTConnectionString"].ConnectionString;
        string constrMVint = ConfigurationManager.ConnectionStrings["MVIntConnectionString"].ConnectionString;

        public VendorCommitApp()
        {
            loaded = false;

            InitializeComponent();
            MyDataGridViewInitializationMethod();
            // setSpreadsheetListView();   
            getauthorization();
            vendorCodeVendorMater();
            loadVendorCodeE2Open();
            loadShareDoc();
            loadWeekYearCombo();
            loadWeekYear();
            getCarrier();
            loadVendorCode();
            loadVendorMasterGrid();
            loadStatus();
            loadFOB();
            loadShipMethod();
            loadPlanType();
            this.grdVendPart.DataSource = null;
        }

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


        public bool validatePart(string part_num)
        {
            string qry = "select * from inv_master where part_no='" + part_num + "'";
            if (getdataSet(qry).Tables[0].Rows.Count > 0)
                return true;
            else
                return false;
        }

        public void loadWeekYearCombo()
        {

            string sqlstr = "   SELECT distinct mweek AS week_desc,mweek AS week_value  FROM  dbo.m_WeekNumber select [dbo].[udf_GetISOWeekNumberFromDate](getdate())  select year(getdate()) ";
            DataSet ds = getData(sqlstr);


            DataTable dtYear = new DataTable();
            dtYear.Columns.Add("year");
            DataRow dr = dtYear.NewRow();
            dr[0] = "2015";
            dtYear.Rows.Add(dr);
            dr = dtYear.NewRow();
            dr[0] = "2016";
            dtYear.Rows.Add(dr);
            dr = dtYear.NewRow();
            dr[0] = "2017";
            dtYear.Rows.Add(dr);
            dr = dtYear.NewRow();
            dr[0] = "2018";
            dtYear.Rows.Add(dr);
            dr = dtYear.NewRow();
            dr[0] = "2019";
            dtYear.Rows.Add(dr);
            dr = dtYear.NewRow();
            dr[0] = "2020";
            dtYear.Rows.Add(dr);

            cmbYearRequest.DataSource = dtYear;
            cmbYearRequest.DisplayMember = "year";
            cmbYearRequest.ValueMember = "year";
            cmbYearRequest.SelectedValue = ds.Tables[2].Rows[0][0].ToString();

            cmbWeekRequest.DataSource = ds.Tables[0];
            cmbWeekRequest.DisplayMember = "week_desc";
            cmbWeekRequest.ValueMember = "week_desc";
            cmbWeekRequest.SelectedValue = ds.Tables[1].Rows[0][0].ToString();

        }

        public void getRequestId()
        {
            try
            {
                if (cmbYearRequest.Text.Trim() != "" && cmbWeekRequest.Text.Trim() != "" && this.cmbVendorRequest.SelectedValue.ToString() != "0")
                {
                    string req = cmbYearRequest.Text.Trim().Substring(2, 2) + cmbWeekRequest.Text.ToString().PadLeft(2, '0') + cmbVendorRequest.SelectedValue.ToString();
                    txtReqID.Value = Convert.ToInt32(req);
                }
            }
            catch { }
        }

        public void loadVendorCodeE2Open()
        {

            string sqlstr = " SELECT distinct  [E2Open_VendID],VC_EPVendCode   FROM [MIMDIST].[dbo].[VCAPP_VendPart] where isnull(E2Open_vendID,'')!=''";
            DataSet ds = getData(sqlstr);
            DataTable dtVendor = ds.Tables[0].Copy();
            DataRow dr = dtVendor.NewRow();
            dr["VC_EPVendCode"] = "SELECT";
            dr["E2Open_VendID"] = "";
            dtVendor.Rows.InsertAt(dr, 0);

          

            cmbVendorE2Open.DataSource = dtVendor;
            cmbVendorE2Open.DisplayMember = "VC_EPVendCode";
            cmbVendorE2Open.ValueMember = "E2Open_VendID";
            cmbVendorE2Open.SelectedIndex = 0;

        }

        public void loadVendorCode()
        {
            string sqlstr = " SELECT vend_id, ep_vend_code, vend_type, shared_doc_name  FROM [MIMDIST].[dbo].[vc_vend_master] where active='Y'  order by ep_vend_code ";
            DataSet ds = getData(sqlstr);
            DataTable dtVendor = ds.Tables[0].Copy();
            DataRow dr = ds.Tables[0].NewRow();
            dr["ep_vend_code"] = "SELECT";
            dr["vend_id"] = "0";
            ds.Tables[0].Rows.InsertAt(dr, 0);

            // DataTable dtVendor = ds.Tables[0].Copy();
            DataTable VendorLoadShip = ds.Tables[0].Copy();
            DataTable VendorRequest = ds.Tables[0].Copy();
            DataTable VendPart = ds.Tables[0].Copy();

            DataRow drvend = dtVendor.NewRow();
            drvend["ep_vend_code"] = "ALL";
            drvend["vend_id"] = "0";
            dtVendor.Rows.InsertAt(drvend, 0);

            cmbVendor.DataSource = dtVendor;
            cmbVendor.DisplayMember = "ep_vend_code";
            cmbVendor.ValueMember = "ep_vend_code";
            cmbVendor.SelectedIndex = 0;

            this.cmbVendorLoadShip.DataSource = VendorLoadShip;
            cmbVendorLoadShip.DisplayMember = "ep_vend_code";
            cmbVendorLoadShip.ValueMember = "vend_id";
            cmbVendorLoadShip.SelectedIndex = 0;

            cmbVendorRequest.DataSource = VendorRequest;
            cmbVendorRequest.DisplayMember = "ep_vend_code";
            cmbVendorRequest.ValueMember = "vend_id";
            cmbVendorRequest.SelectedIndex = 0;

            cmbVendPart.DataSource = VendPart;
            cmbVendPart.DisplayMember = "ep_vend_code";
            cmbVendPart.ValueMember = "vend_id";
            cmbVendPart.SelectedIndex = 0;

        

        }


        public void loadWeekYear()
        {
            string sqlstr = "   SELECT DISTINCT TOP (100) PERCENT dbo.m_WeekNumber.yearweek AS value, dbo.m_WeekNumber.yearweek AS descr  FROM    dbo.m_WeekNumber INNER JOIN  dbo.VCA_ReqCom_Table ON dbo.m_WeekNumber.yearweek = dbo.VCA_ReqCom_Table.REQ_YRWK WHERE     (dbo.m_WeekNumber.StartDate > GETDATE() - 14)  ORDER BY value";//union select cast(year(getdate()) as varchar) as descr,year(getdate()) as value select year(getdate())  ";
            DataSet ds = getData(sqlstr);
            DataRow dr = ds.Tables[0].NewRow();
            dr["descr"] = "ALL";
            dr["value"] = "";
            ds.Tables[0].Rows.Add(dr);
            cmbYearWeek2.DataSource = ds.Tables[0];
            cmbYearWeek2.DisplayMember = "descr";
            cmbYearWeek2.ValueMember = "value";
            cmbYearWeek2.SelectedIndex = cmbYearWeek2.Items.Count - 1;

            sqlstr = "   SELECT distinct case when SHP_Recvd='Y' then 'YES' else 'NO' end as descr,SHP_Recvd as value FROM  VCA_ShipTrack_Table";
            ds = getData(sqlstr);
            DataRow dr2 = ds.Tables[0].NewRow();
            dr2["descr"] = "ALL";
            dr2["value"] = "";
            ds.Tables[0].Rows.Add(dr2);
            cmbReceived.DataSource = ds.Tables[0];
            cmbReceived.DisplayMember = "descr";
            cmbReceived.ValueMember = "value";
            cmbReceived.Text = "NO";

        }


        private class Item
        {
            public string Name;
            public string Value;
            public Item(string name, string value)
            {
                Name = name; Value = value;
            }

            public override string ToString()
            {
                // Generates the text shown in the combo box
                return Name;
            }
        }

        string filename;
        DriveService service;

        public void getauthorization()
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

            /*   string pageToken = null;
               string fileId = "";

            do
               {
                   var request = service.Files.List();
                   // request.Q = "mimeType = 'application/vnd.google-apps.folder' ";
                   string parent_id = "0B1auUUiGcbbBVm9OQ25uYjhrYXc";
                   request.Q = "'" + parent_id + "' in parents";
                   request.Spaces = "Drive";
                   request.Fields = "nextPageToken, files(id, name)";
                   request.PageToken = pageToken;

                   var result = request.Execute();
                   foreach (var file in result.Files)
                   {
                   
                       if (file.Name.StartsWith("Demand Forecast"))
                       {
                          ListViewItem item = new ListViewItem(new string[2] { file.Name, file.Id });
                              // googleshareDoc = file1.Title;
                               this.lstWorksheet.Items.Add(item);
                          // fileId = file.Id;
                           //break;

                       }
                   }
                   pageToken = result.NextPageToken;
               } while (pageToken != null);

               string fileName = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\currentRackDeliveryPlan.xlsx";
               fileName = fileName.Replace("file:\\", "");

               Google.Apis.Drive.v3.Data.File file1 = service.Files.Get(fileId).Execute();
            //   if (this.downloadfile(service, fileId, fileName))
             //  {

               //    Demand ForecastgetallSheet(fileName);
                   //  MessageBox.Show("File has been downloaded");

               //}
              */
        }


        void getCollection(string folderName, string filename, string FirstorALL)
        {

            /*    UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                      new ClientSecrets
                      {
                          ClientId = "679763664590-5tugcvfsp90pbc707t4f8dlma3oj9bab.apps.googleusercontent.com",
                          ClientSecret = "2lneARBZtpxrbX9vVkxcDiXe",
                      },
                      new[] { DriveService.Scope.Drive },
                      "user", CancellationToken.None).Result;
                service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "MIM Project ",
                });
            
            
            string googleshareDoc = "";
          //  DataTable dtShareDoc = loadShareDoc();
          //  int doccount = dtShareDoc.Rows.Count;
         //   int docIndex = 0;
           // string docName = "";
        //    docName = dtShareDoc.Rows[docIndex][0].ToString();
            string folder_id = "";
            try
            {
           
        
            
            FilesResource.ListRequest lstreq = service.Files.List();
                do
                {
                    try
                    {
                        lstreq.Q = " title ='" + folderName + "'";
                        FileList fileList = lstreq.Execute();
                        foreach (Google.Apis.Drive.v2.Data.File fileitem in fileList.Items)
                        {
                                ListViewItem item = new ListViewItem(new string[2] { fileitem.Title, fileitem.Id });
                                googleshareDoc = fileitem.Title;
                                folder_id = fileitem.Id;
                                goto  findChild;                               
                              //  this.spreadsheetListView.Items.Add(item);                          
                        }

                        lstreq.PageToken = fileList.NextPageToken;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("An error occurred: " + e.Message);
                        lstreq.PageToken = null;
                    }
                } while (!String.IsNullOrEmpty(lstreq.PageToken));
              
            }
            catch (Exception exp)
            {
                MessageBox.Show("Exception" + exp.Message.ToString());
            }
            findChild:
                getALLchildInFolder(service, folder_id, filename, FirstorALL);       */

        }



        public DataTable loadShareDoc()
        {
            string sqlstr = " SELECT distinct [shared_doc_name]   FROM [MIMDIST].[dbo].[vc_vend_master] where active ='Y' order by [shared_doc_name]  ";
            DataTable dt = getData(sqlstr).Tables[0];

            foreach (DataRow dr in dt.Rows)
            {
                ListViewItem item = new ListViewItem(new string[2] { dr["shared_doc_name"].ToString(), dr["shared_doc_name"].ToString() });

                this.spreadsheetListView.Items.Add(item);
            }
            return dt;
        }

        private void listView1_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            // If this item is the selected item
            if (e.Item.Selected)
            {
                // If the selected item just lost the focus
                /*******************   if (gListView1LostFocusItem == e.Item.Index)
                   {
                       // Set the colors to whatever you want (I would suggest
                       // something less intense than the colors used for the
                       // selected item when it has focus)
                       e.Item.ForeColor = Color.Black;
                       e.Item.BackColor = Color.LightBlue;

                       // Indicate that this action does not need to be performed
                       // again (until the next time the selected item loses focus)
                       gListView1LostFocusItem = -1;
                   }
                   else if (spreadsheetListView.Focused)  // If the selected item has focus
                   {
                       // Set the colors to the normal colors for a selected item
                       e.Item.ForeColor = SystemColors.HighlightText;
                       e.Item.BackColor = SystemColors.Highlight;
                   }**********************/
            }
            else
            {
                // Set the normal colors for items that are not selected
                e.Item.ForeColor = spreadsheetListView.ForeColor;
                e.Item.BackColor = spreadsheetListView.BackColor;
            }

            e.DrawBackground();
            e.DrawText();
        }


        private void worksheetListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (worksheetListView.SelectedItems.Count > 0)
            //{
            //    // MessageBox.Show(worksheetListView.SelectedItems[0].SubItems[2].Text);
            //    ListQuery query = new ListQuery(worksheetListView.SelectedItems[0].SubItems[3].Text);
            //    SetListListView(this.spreadsheetService.Query(query));
            //}
        }
        DataTable table = new DataTable();

        bool loaded = false;

        public DataSet getDataSet(string sqlStr)
        {
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlStr, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;
        }
        //void SetListListView(ListFeed feed)
        //{
        //    SqlConnection cn = new SqlConnection(constr);
        //    SqlCommand cmd = new SqlCommand("select * from m_temp_build_plan");

        //    AtomEntryCollection entries = feed.Entries;
        //    table = new DataTable();




        //    //  MessageBox.Show(entries.Count.ToString());
        //    for (int i = 0; i < entries.Count; i++)
        //    {
        //        ListEntry entry = entries[i] as ListEntry;
        //        ListViewItem item = new ListViewItem();
        //        if (entry != null)
        //        {
        //            ListEntry.CustomElementCollection elements = entry.Elements;

        //            DataRow myDataRow = table.NewRow();



        //            for (int j = 0; j < elements.Count; j++)
        //            {
        //                if (j > table.Columns.Count - 1)
        //                {
        //                    table.Columns.Add("col" + j);
        //                }

        //                if (elements[j].Value != null || elements[j].Value != "0" || elements[j].Value != "")
        //                {
        //                    myDataRow[j] = elements[j].Value;
        //                }
        //                else
        //                {
        //                    myDataRow[j] = "";
        //                }
        //            }
        //            table.Rows.Add(myDataRow);
        //        }
        //    }
        //    this.dataGridView1.DataSource = table;

        //    Cursor.Current = Cursors.Default;

        //}

        DataTable tableGrid;
        string sheetName = "";
        public static bool IsDate(Object obj)
        {
            string strDate = obj.ToString();
            try
            {
                DateTime dt = DateTime.Parse(strDate);
                // if ((dt.Month <13 ) || (dt.Day < 1 && dt.Day > 31) || dt.Year != System.DateTime.Now.Year)
                //     return false;
                // else
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void openExcelInSpreadSheet(string filename, FarPoint.Win.Spread.FpSpread FP)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = null;
            Microsoft.Office.Interop.Excel.Sheets excelSheets = null;
            //Microsoft.Office.Interop.Excel.Application excelApp;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook1 = null;
            Microsoft.Office.Interop.Excel.Workbooks excelWorkbooks1 = null;
            Microsoft.Office.Interop.Excel.Application excelApp1 = new Microsoft.Office.Interop.Excel.Application();

            string workbookPath = filename.Replace("file:\\", "");
            excelWorkbooks1 = excelApp1.Workbooks;
            excelWorkbook1 = excelWorkbooks1.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            excelWorkbook1.Save();
            excelWorkbook1.Close(true);
            try
            {
                //  FarPoint.Win.Spread.FpSpread fp = new FarPoint.Win.Spread.FpSpread();
                // fpSpread1.OpenExcel(workbookPath);
                FP.OpenExcel(workbookPath);
            }
            catch (Exception ex)
            { }
        }

        void SetWorksheetListView()
        {


            Microsoft.Office.Interop.Excel.Worksheet sheet = null;
            Microsoft.Office.Interop.Excel.Sheets excelSheets = null;
            //Microsoft.Office.Interop.Excel.Application excelApp;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            Microsoft.Office.Interop.Excel.Workbooks excelWorkbooks = null;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            string workbookPath = saveTo.Replace("file:\\", "");
            excelWorkbooks = excelApp.Workbooks;
            excelWorkbook = excelWorkbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            excelWorkbook.Close(true);
            try
            {
                //  FarPoint.Win.Spread.FpSpread fp = new FarPoint.Win.Spread.FpSpread();
                fpSpread1.OpenExcel(workbookPath);
            }
            catch (Exception ex)
            { }
            //this.tabSaveVendorCommit.Controls.Add(fp);
            //fp.Location = new Point(0, 40);
            //tabControl1.TabPages[2].Controls.Add(fp);
            /*       int i=0;
                   try
                   {

                       excelWorkbooks = excelApp.Workbooks;
                       excelWorkbook = excelWorkbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


                       excelSheets = excelWorkbook.Worksheets;
                       string worksheetname = "";
                       sheet = excelSheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
      
                       DataRow dr = tableGrid.NewRow();
                       int curr_row = 0;
                       int last_row = 0;
                       int rowNot = 0;

                       int first_part_row = 5;//6
                       int num_of_row_in_part = 12;// 14

                       //   curr_row = 0;
                       //   last_row = 0;

                       rowNot = 0;
                       int total_mim_req = (first_part_row + 3);
                       int total_vend_com = (first_part_row + 7);
                       int ship_row = first_part_row + 8;
                       int ship_end_row = first_part_row + 12;
                       dr = tableGrid.NewRow();
                       String create_date = (sheet.Cells[1, 2] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                       String RequestID = (sheet.Cells[3, 2] as Microsoft.Office.Interop.Excel.Range).Value.ToString();

                       if (IsDate(create_date))
                       {
                           dr[0] = (DateTime.Parse(create_date)).Date.ToString("d");
                       }
                       else
                       {
                           dr[0] = create_date;
                       }

                       tableGrid.Rows.Add(dr);


                       string curr_quarter = (string)(sheet.Cells[3, 2] as Microsoft.Office.Interop.Excel.Range).Value.ToString();

                       for ( i = 2; i < 4; i++)
                       {
                           dr = tableGrid.NewRow();
                           if (i != 3)
                               dr[0] = (sheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value.ToString();                   
                               dr[1] = (sheet.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                           for (int j = 6; j <= 42; j++)
                           {
                               if (i == 3)
                               {
                                   String dt = "";
                                   if ((sheet.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value != null)
                                   {
                                       dt = (sheet.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                                   }
                                   if (IsDate(dt))
                                   {
                                       dr[j - 5] = (DateTime.Parse(dt)).Date.ToString("d");
                                   }
                                   else
                                   {
                                       dr[j - 5] = dt;
                                   }

                                   //  dr[j - 5] = dt;

                               }
                               else
                               {
                                   dr[j - 5] = "";
                                   try
                                   {
                                       dr[j - 5] = (sheet.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                                   }
                                   catch
                                   { }
                               }
                           }
                           tableGrid.Rows.Add(dr);
                       }
               //    }
                //   catch (Exception e)
               //    {
                       //MessageBox.Show(e.Message.ToString());
                //   }


                       dr = tableGrid.NewRow();

                       for (int cnt = 0; cnt < tableGrid.Columns.Count; cnt++)
                           dr[cnt] = "";

                       tableGrid.Rows.Add(dr);

                       int rowCount = sheet.Rows.Count;
                       for ( i = 5; i <= rowCount; i++)
                       {
                           string chkfirst = "";
                           string chksec = "";
                           string chkthird = "";
                           try
                           {
                               try
                               {
                                   chkfirst = (sheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                               }
                               catch
                               { }

                               try
                               {
                                   chksec = (sheet.Cells[i + 1, 1] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                               }
                               catch { }

                               try
                               {
                                   chkthird = (sheet.Cells[i + 1, 1] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                               }
                               catch { }
                               if ((chkfirst == "") && (chksec == "") && (chkthird == ""))
                                   break;
                           }
                           catch
                           {
                               break;
                           }

                           if (i == ship_end_row)
                           {
                               first_part_row = first_part_row + num_of_row_in_part;
                               total_mim_req = (first_part_row + 3);
                               total_vend_com = (first_part_row + 7);
                               ship_row = first_part_row + 8;
                               ship_end_row = first_part_row + 12;


                          //     int total_mim_req = (first_part_row + 3);
                           //    int total_vend_com = (first_part_row + 7);
                           //    int ship_row = first_part_row + 8;
                           //    int ship_end_row = first_part_row + 12;

                           }
                           if ((i < ship_row || i > ship_end_row) && (i != total_mim_req))
                           {
                               if (i > 5) //&& entry.Cell.Row <= ship_row
                               {
                                   if (last_row == 0)
                                       last_row = i;
                                   curr_row = i;

                                   if (last_row != curr_row)
                                   {
                                       last_row = i;
                                       //tableGrid.Rows.Add(dr);
                                       dr = tableGrid.NewRow();
                                       if (last_row == 6)
                                       {
                                           //  dr = tableGrid.NewRow();
                                           for (int cnt = 0; cnt < tableGrid.Columns.Count; cnt++)
                                               dr[cnt] = "";
                                           tableGrid.Rows.Add(dr);
                                           dr = tableGrid.NewRow();
                                       }
                                   }
                               }

                               dr = tableGrid.NewRow();
                               dr[0] = (sheet.Cells[first_part_row, 2] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                               dr[1] = (sheet.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                               for (int c = 6; c <= 41; c++)
                               {
                                   try
                                   {
                                       if ((sheet.Cells[i, c] as Microsoft.Office.Interop.Excel.Range).Value.ToString().Trim() == "-")
                                           dr[(int)(c - 5)] = 0;
                                       else
                                       dr[(int)(c - 5)] = (sheet.Cells[i, c] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                                   }
                                   catch {

                                       dr[(int)(c - 5)] = 0;
                                   }
                               }
                               tableGrid.Rows.Add(dr);

                               if (i == total_vend_com)
                               {
                                   for (int k = 0; k < tableGrid.Columns.Count; k++)
                                       dr[k] = "";

                               }
                           }
                       }                      
               
                   }
                   catch (Exception e)
                   {
                        MessageBox.Show(i.ToString());
                   }
                   finally
                   { 
                       //excelWorkbook.Save();
                       excelWorkbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                       excelWorkbooks.Close();
                       excelApp.Quit();
                      // System.Runtime.InteropServices.Marshal.ReleaseComObject(excelUsedRange);
                       System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                       System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheets);
                       System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbooks);
                       System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                       System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                       //excelUsedRange = null;
                       //sheet = null;
                       //excelSheets = null;
                       //excelWorkbooks = null;
                       //excelWorkbook = null;
                       //excelApp = null;

                       GC.GetTotalMemory(false);
                       GC.Collect();
                       GC.WaitForPendingFinalizers();
                       GC.Collect();
                       GC.GetTotalMemory(true);

                      // System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlRng);
                       System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sheet);

                   //      System.Runtime.InteropServices.xlBook.Close(Type.Missing, Type.Missing, Type.Missing);
                       System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbooks);

                      // excelApp.Quit();
                       System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                       System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);           
                       TryToDelete(workbookPath);
                   }
              
                //   FarPoint.Win.Spread.FpSpread fpSpread1 = new FarPoint.Win.Spread.FpSpread();
                  // fpSpread1.DataSource = tableGrid;
                   //tabControl1.TabPages[2].Controls.Add(fpSpread1);
                   fpSpread1.DataSource = tableGrid;
                   dataGridView1.DataSource = tableGrid;
                   for (int c = 0; c < dataGridView1.Columns.Count; c++)
                       dataGridView1.Columns[c].SortMode = DataGridViewColumnSortMode.NotSortable;
                */
        }


        static bool TryToDelete(string f)
        {
            try
            {
                // A.
                // Try to delete the file.
                System.IO.File.Delete(f);
                return true;
            }
            catch (System.IO.IOException exp)
            {
                // B.
                // We could not delete the file.
                string msg = exp.Message.ToString();
                return false;
            }
        }

        private void lstWorksheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            ListItem1 lst = lstWorksheet1.SelectedItem as ListItem1;
            //WorksheetQuery query = new WorksheetQuery(lst.ID.ToString());
            //SetWorksheetListView(spreadsheetService.Query(query));
            Cursor.Current = Cursors.Default;
        }
        string saveTo = "CurrentVendorCommit";
        /*    public Boolean downloadFile(DriveService _service, File _fileResource, string _saveTo)
            {
                if (!String.IsNullOrEmpty(_fileResource.ExportLinks["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]))
                {
                    try
                    {
                        var x = _service.HttpClient.GetByteArrayAsync(_fileResource.ExportLinks["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]);
                     //   var x = _service.HttpClient.GetByteArrayAsync(_fileResource.ExportLinks["text/csv"]);
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
            }
        */
        int sheetnum = 0;
        string filename1 = "";


        private void spreadsheetListView_SelectedIndexChanged(object sender, EventArgs e)
        {

            Cursor.Current = Cursors.WaitCursor;

            this.lstWorksheet.Items.Clear();

            if (spreadsheetListView.SelectedItems.Count > 0)
            {
                string pageToken = null;
                do
                {

                    string parent_id = getGIDByName(spreadsheetListView.SelectedItems[0].SubItems[1].Text);
                    var result = getFilesByParentId(parent_id, pageToken);
                    foreach (var file in result.Files)
                    {
                        if (file.Name.StartsWith("Demand Forecast"))
                        {
                            ListViewItem item = new ListViewItem(new string[2] { file.Name, file.Id });
                            // googleshareDoc = file1.Title;
                            this.lstWorksheet.Items.Add(item);
                            // fileId = file.Id;
                            //break;
                        }
                    }
                    pageToken = result.NextPageToken;
                } while (pageToken != null);

                // getCollection(spreadsheetListView.SelectedItems[0].SubItems[1].Text, "Demand Forecast","ALL");  


            }
            /*  if (spreadsheetListView.SelectedItems.Count > 0)
              {

                  filename = spreadsheetListView.SelectedItems[0].SubItems[1].Text;
                  //  googleshareDoc = spreadsheetListView.;
                  //  fileId = spreadsheetListView.SelectedItems[0]..Text;
                  Google.Apis.Drive.v2.Data.File file1 = service.Files.Get(filename).Execute();
                  saveTo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\currentForecast.xlsx";
                  saveTo = saveTo.Replace("file:\\", "");
                  if (downloadFile(service, file1, saveTo))
                      getallSheet(saveTo);
              }*/


            Cursor.Current = Cursors.Default;
        }

        public FileList getFilesByParentId(string parent_id, string pageToken)
        {
            // request.Q = "mimeType = 'application/vnd.google-apps.folder' ";

            var request = service.Files.List();
            request.Q = "'" + parent_id + "' in parents";
            request.Spaces = "Drive";
            request.Fields = "nextPageToken, files(id, name)";
            request.PageToken = pageToken;

            var result = request.Execute();
            return result;
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


        public void showInGrid()
        {

            //  Microsoft.Office.Interop.Excel.Worksheet sheet = null;
            Microsoft.Office.Interop.Excel.Sheets excelSheets = null;
            //Microsoft.Office.Interop.Excel.Application excelApp;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            // Microsoft.Office.Interop.Excel.Workbooks excelWorkbooks = null;

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            string workbookPath = saveTo.Replace("file:\\", "");
            excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            sheetnum = 0;
            excelSheets = excelWorkbook.Worksheets;
            string worksheetname = "";
            foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in excelWorkbook.Worksheets)
            {
                worksheetname = worksheet.Name;
                break;

            }
            /*   System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + saveTo + "';Extended Properties=Excel 12.0;");
               System.Data.OleDb.OleDbDataAdapter MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [" +worksheetname +"$]", MyConnection);
               MyCommand.TableMappings.Add("Table", "TestTable");
               System.Data.DataSet DtSet = new System.Data.DataSet();
               MyCommand.Fill(DtSet);
               DataTable myTable = DtSet.Tables[0];*/






            //   dataGridView1.DataSource = myTable;

            //   MyCommand.Fill(DtSet);
            //   dataGridView1.DataSource = DtSet.Tables[0];
            //  lblNumOfRow.Text = DtSet.Tables[0].Rows.Count.ToString();
            //  MyConnection.Close();


        }

        string googleshareDoc = "";
        string fileLink;
        /*   public void  getALLchildInFolder(DriveService service, String folderId,string filestart,string FirstorALL)
           {
               this.lstWorksheet.Items.Clear();
               ChildrenResource.ListRequest request = service.Children.List(folderId);
           
               do
               {
                   try
                   {
                
                       request.Q = " title  contains '"+filestart +"'";
                       ChildList children = request.Execute();

                       foreach (ChildReference child in children.Items)
                       {
                           Google.Apis.Drive.v2.Data.File file1 = service.Files.Get(child.Id).Execute();                       
                         //  if (file1.Title.Contains("Forecast Wk"))
                         //  {
                           if (FirstorALL == "First")
                           {
                               fileLink = file1.Id;
                               goto endogproc;
                           }
                           else
                           {
                               ListViewItem item = new ListViewItem(new string[2] { file1.Title, child.Id });
                               googleshareDoc = file1.Title;
                               this.lstWorksheet.Items.Add(item);
                           }
                          // }
                       }

                       request.PageToken = children.NextPageToken;
                   }
                   catch (Exception e)
                   {
                       Console.WriteLine("An error occurred: " + e.Message);
                       request.PageToken = null;
                   }
               } while (!String.IsNullOrEmpty(request.PageToken));
           endogproc:
               fileLink = fileLink;
        
           }

           public string  getFirstchildInFolder(DriveService service, String folderId, string filestart)
           {
               this.lstWorksheet.Items.Clear();
               ChildrenResource.ListRequest request = service.Children.List(folderId);
               string fileLink = "";
               do
               {
                   try
                   {
                       request.Q = " title  contains '" + filestart + "'";
                       ChildList children = request.Execute();
                       foreach (ChildReference child in children.Items)
                       {
                           Google.Apis.Drive.v2.Data.File file1 = service.Files.Get(child.Id).Execute();
                           //  if (file1.Title.Contains("Forecast Wk"))
                           //  {
                           fileLink = file1.Id;
                           goto endogproc;
                           // }
                       }

                       request.PageToken = children.NextPageToken;
                   }
                   catch (Exception e)
                   {
                       Console.WriteLine("An error occurred: " + e.Message);
                       request.PageToken = null;
                   }
               } while (!String.IsNullOrEmpty(request.PageToken));
           endogproc:
               return fileLink;
           }

           */

        private void btnDownloadGrid_Click(object sender, EventArgs e)
        {

        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            /*    string strheader=" insert into vca_req_com(vendor,req_doc_date,gpn,req_YearWeek,req_date,req_site,req_qty,com_qty) values ";
                string v_req_doc_date=tableGrid.Rows[0][0].ToString();
                string v_vendor=tableGrid.Rows[1][0].ToString();

                string[] arrYear=new string[36];
                string[] arrWeek = new string[36];
                string[] arrYearWeek = new string[36];
                string[] arrDate = new string[36];
                string[] at_req_qty = new string[36];
                string[] at_com_qty = new string[36];
                string[] nl_req_qty = new string[36];
                string[] nl_com_qty = new string[36];
                string[] sg_req_qty = new string[36];
                string[] sg_com_qty = new string[36];
                string str="";
                string strAT = "";
                string strNL = "";
                string strSG = "";
                int first_part_row_num1 = 5;
                int first_part_row_num = 0;
                int part_num_gap = 7;
                int part_count = (tableGrid.Rows.Count - 3) / part_num_gap;
                for (int num_of_part = 0; num_of_part < part_count; num_of_part++)
                {
                    first_part_row_num = first_part_row_num1 + part_num_gap * num_of_part;
                    for (int j=2;j<38;j++)
                    {
                        arrYear[j-2]= tableGrid.Rows[1][j].ToString();
                        arrWeek[j-2]= tableGrid.Rows[2][j].ToString();
                        if (tableGrid.Rows[2][j].ToString().Length==1)
                           arrYearWeek[j - 2] = tableGrid.Rows[1][j].ToString() + "W0" + tableGrid.Rows[2][j].ToString();
                        else if (tableGrid.Rows[2][j].ToString().Length>1)
                            arrYearWeek[j - 2] = tableGrid.Rows[1][j].ToString() + "W" + tableGrid.Rows[2][j].ToString();
                        arrDate[j-2]= tableGrid.Rows[3][j].ToString();
                        if (tableGrid.Rows[first_part_row_num][j].ToString() == "")
                            at_req_qty[j - 2] = "0";
                        else
                            at_req_qty[j - 2] = tableGrid.Rows[first_part_row_num][j].ToString().Replace(",","");


                        if (tableGrid.Rows[first_part_row_num + 1][j].ToString() == "")// || tableGrid.Rows[first_part_row_num + 1][j].ToString().Trim() == "-")
                            nl_req_qty[j - 2] = "0";
                        else
                            nl_req_qty[j - 2] = tableGrid.Rows[first_part_row_num + 1][j].ToString().Replace(",", "");

                        if (tableGrid.Rows[first_part_row_num + 2][j].ToString() == "")// || tableGrid.Rows[first_part_row_num + 2][j].ToString().Trim() == "-")
                            sg_req_qty[j - 2] = "0";
                        else
                            sg_req_qty[j - 2] = tableGrid.Rows[first_part_row_num + 2][j].ToString().Replace(",", "");

                        if (tableGrid.Rows[first_part_row_num + 3][j].ToString() == "")// || tableGrid.Rows[first_part_row_num + 3][j].ToString().Trim() == "-")
                            at_com_qty[j - 2] = "0";
                        else
                            at_com_qty[j - 2] = tableGrid.Rows[first_part_row_num + 3][j].ToString().Replace(",", "");

                        if (tableGrid.Rows[first_part_row_num + 4][j].ToString() == "")// || tableGrid.Rows[first_part_row_num + 4][j].ToString().Trim() == "-")
                            nl_com_qty[j - 2] = "0";
                        else
                            nl_com_qty[j - 2] = tableGrid.Rows[first_part_row_num + 4][j].ToString().Replace(",", "");

                        if (tableGrid.Rows[first_part_row_num + 5][j].ToString() == "")// || tableGrid.Rows[first_part_row_num + 5][j].ToString().Trim() == "-")
                            sg_com_qty[j - 2] = "0";
                        else
                            sg_com_qty[j - 2] = tableGrid.Rows[first_part_row_num + 5][j].ToString().Replace(",", "");  
                    }

                    string v_gpn = tableGrid.Rows[first_part_row_num][0].ToString();
                    for (int col = 0; col < 36; col++)
                    {
                        if (at_req_qty[col].ToString() != "0" || at_com_qty[col] != "0")
                            strAT = strheader + "( '" + v_vendor + "','" + v_req_doc_date + "','" + v_gpn + "','" + arrYearWeek[col] + "','" + arrDate[col] + "','AT','" + at_req_qty[col] + "','" + at_com_qty[col] + "')";
                        else
                            strAT = "";
                        if (nl_req_qty[col] != "0" || nl_com_qty[col] != "0" )
                            strNL = strheader + "( '" + v_vendor + "','" + v_req_doc_date + "','" + v_gpn + "','" + arrYearWeek[col] + "','" + arrDate[col] + "','NL','" + nl_req_qty[col] + "','" + nl_com_qty[col] + "')";
                        else
                            strNL = "";
                        if (sg_req_qty[col] != "0" || sg_com_qty[col] != "0")
                            strSG = strheader + "( '" + v_vendor + "','" + v_req_doc_date + "','" + v_gpn + "','" + arrYearWeek[col] + "','" + arrDate[col] + "','SG','" + sg_req_qty[col] + "','" + sg_com_qty[col] + "')";
                        else
                            strSG = "";
                        str = str + strAT + strNL + strSG;
                    }
                }

             str = str + " exec m_updateCommitID " ; //";
                try
                {
                    SqlConnection cn = new SqlConnection(constr);
                    cn.Open();
                    SqlCommand cmd = new SqlCommand(str, cn);
                    cmd.CommandTimeout = 0;
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Vendor Commit has been saved successfully");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error to save Vendor Commit! Please check"+ ex.Message.ToString());
                }*/
        }



        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public DataSet getData(string sqlstr)
        {
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlstr, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet dsVendor = new DataSet();
            da.Fill(dsVendor);
            cn.Dispose();
            cn.Close();
            return dsVendor;
        }

        DataTable dtShipTracker;
        DataTable dtdbShipTracker;

        private void btnShow_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            loadDatagrid();
            Cursor.Current = Cursors.Default;
        }


        public void loadDatagrid()
        {
            /*     string strVendor = " SELECT    ID,vendor as VENDOR,ship_id as [SHIP ID],received as [RECEIVED],eta_date as [ETA DATE],po_num as [PO NUM],gpn as [GPN],ship_to as [SHIP TO],ship_date as [SHIP DATE],ship_qty as [SHIP QTY], packing_slip as [PACKING SLIP], ship_method as [SHIP METHOD],carrier as [CARRIER], tracking_no as [TRACKING NUM] FROM         vca_ship_tracker_for_vcapp  where vendor<>''";
        
                 if (cmbVendor.SelectedValue != "")
                     strVendor = strVendor + " and vendor='" + cmbVendor.SelectedValue + "'";

                 if (cmbWeek.SelectedValue.ToString() != "0")
                     strVendor = strVendor + " and req_week='" + cmbWeek.SelectedValue + "'";

                 if (cmbYear.SelectedValue.ToString() != "0")
                     strVendor = strVendor + " and req_year='" + cmbYear.SelectedValue + "'";

                 if (cmbReceived.SelectedValue != "")
                     strVendor = strVendor + " and received='" + cmbReceived.SelectedValue + "'";
                 if (txtPartNum.Text.ToString().Trim() != "")
                     strVendor = strVendor + " and gpn like '" + txtPartNum.Text.Trim() + "%'";


                 SqlConnection cn = new SqlConnection(constr);
                 cn.Open();
                 SqlCommand cmd = new SqlCommand(strVendor, cn);
                 SqlDataAdapter da = new SqlDataAdapter(cmd);
                 DataSet dsVendor = new DataSet();
                 da.Fill(dsVendor);
                 cn.Dispose();
                 cn.Close();
                 ultraGrid1.DataSource = dsVendor.Tables[0];
                 */
            dgvShipTracker.EditMode = DataGridViewEditMode.EditOnKeystroke;
            dgvShipTracker.Columns.Clear();
            // string strVendor = " SELECT    ID,vendor as VENDOR,ship_id as [SHIP ID],[RECEIVED],eta_date as [ETA DATE],po_num as [PO NUM],gpn as [GPN],ship_to as [SHIP TO],ship_date as [SHIP DATE],ship_qty as [SHIP QTY], packing_slip as [PACKING SLIP], ship_method as [SHIP METHOD],carrier as [CARRIER], tracking_no as [TRACKING NUM] FROM         vca_ship_tracker_for_vcapp  where vendor<>''";
            string strVendor = " SELECT    ID,[DOC ID],vendor as VENDOR, [SHIP ID],[RECEIVED],[GPN],[SITE CODE], [SHIP DATE], [SHIP QTY],  [PACKING SLIP],  [SHIP METHOD], [CARRIER],  [TRACKING NUM] FROM         VCA_SHIP_TRACKER_APP where vendor<>''";

            // string strVendor = "SELECT    ID,[DOC ID],vendor as VENDOR, [SHIP ID],[RECEIVED],[GPN],[SITE CODE], [SHIP DATE],substring( (cast([dbo].[udf_GetISOYearFromDate]( [SHIP DATE]) as varchar)+cast([dbo].[udf_GetISOWeekNumberFromDate]( [SHIP DATE]) as varchar)),3,4) [SHIP WK YR], [SHIP QTY],  [PACKING SLIP],  [SHIP METHOD], [CARRIER],  [TRACKING NUM] FROM    VCA_SHIP_TRACKER_APP where vendor<>''";
            if (cmbVendor.Text.ToString() != "ALL")
            {
                strVendor = strVendor + " and vendor='" + cmbVendor.SelectedValue + "'";
            }

            if (cmbWeekYear.Text.Trim() != "")
                strVendor = strVendor + " and REQ_YRWK='" + cmbWeekYear.Text.Trim() + "'";


            if (cmbReceived.SelectedValue.ToString() != "")
                strVendor = strVendor + " and received='" + cmbReceived.SelectedValue + "'";
            if (txtPartNum.Text.ToString().Trim() != "")
                strVendor = strVendor + " and gpn like '" + txtPartNum.Text.Trim() + "%'";

            strVendor = strVendor + "   order by [SHIP DATE]  ";
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(strVendor, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet dsVendor = new DataSet();
            da.Fill(dsVendor);
            cn.Dispose();
            cn.Close();
            dtShipTracker = dsVendor.Tables[0];


            // dtShipTracker.Columns.Add(new DataColumn("ReceivedChk", typeof(bool), "received ='Y'"));

            dgvShipTracker.DataSource = dsVendor.Tables[0];
            dgvShipTracker.AutoGenerateColumns = true;
            dgvShipTracker.Columns[0].Visible = false;

            DataGridViewCheckBoxColumn dgvbc = new DataGridViewCheckBoxColumn();
            dgvbc.DataPropertyName = "RECEIVED";
            dgvbc.Name = "Received_check";
            dgvbc.HeaderText = "RECEIVED";
            dgvbc.TrueValue = "Y";
            dgvbc.FalseValue = "N";


            this.dgvShipTracker.Columns.Insert(3, dgvbc);
            this.dgvShipTracker.CellFormatting += new DataGridViewCellFormattingEventHandler(dataGridView1_CellFormatting);
            DataGridViewLinkColumn links = new DataGridViewLinkColumn();
            links.HeaderText = "TRACKING INFO";
            links.UseColumnTextForLinkValue = true;
            links.Text = "TRACK";
            // links.hr
            links.ActiveLinkColor = Color.White;
            links.LinkBehavior = LinkBehavior.SystemDefault;
            links.LinkColor = Color.Blue;
            links.TrackVisitedState = true;
            links.VisitedLinkColor = Color.YellowGreen;

            dgvShipTracker.Columns.Insert(5, links);
            DataGridViewLinkColumn btnDelete = new DataGridViewLinkColumn();

            btnDelete.HeaderText = "DELETE";
            btnDelete.UseColumnTextForLinkValue = true;
            btnDelete.Text = "Delete";
            // links.hr
            btnDelete.ActiveLinkColor = Color.White;
            btnDelete.LinkBehavior = LinkBehavior.SystemDefault;
            btnDelete.LinkColor = Color.Blue;
            btnDelete.TrackVisitedState = true;
            btnDelete.VisitedLinkColor = Color.YellowGreen;
            dgvShipTracker.Columns[4].Visible = false;
            dgvShipTracker.Columns.Add(btnDelete);

            //for (int col = 0; col < dgvShipTracker.Columns.Count; col++)
            //{
            //    if (col !=3)
            //    dgvShipTracker.Columns[col].ReadOnly = true;
            //}
        }
        void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (e.RowIndex > 7)
                {
                    int cnt = (e.RowIndex - 1) % 7;
                    if (cnt == 0)
                    {

                        if (Convert.ToInt32(dataGridView1.Rows[e.RowIndex - 3].Cells[e.ColumnIndex].Value) > Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
                        {
                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Yellow;
                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.Red;
                        }
                        else if (Convert.ToInt32(dataGridView1.Rows[e.RowIndex - 3].Cells[e.ColumnIndex].Value) < Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
                        {

                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.Red;
                        }

                    }

                    cnt = (e.RowIndex - 2) % 7;
                    if (cnt == 0)
                    {
                        if (Convert.ToInt32(dataGridView1.Rows[e.RowIndex - 3].Cells[e.ColumnIndex].Value) > Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
                        {
                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Yellow;
                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.Red;
                        }
                        else if (Convert.ToInt32(dataGridView1.Rows[e.RowIndex - 3].Cells[e.ColumnIndex].Value) < Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
                        {

                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.Red;
                        }


                    }
                    cnt = (e.RowIndex - 3) % 7;
                    if (cnt == 0)
                    {


                        if (Convert.ToInt32(dataGridView1.Rows[e.RowIndex - 3].Cells[e.ColumnIndex].Value) > Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
                        {
                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Yellow;
                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.Red;
                        }
                        else if (Convert.ToInt32(dataGridView1.Rows[e.RowIndex - 3].Cells[e.ColumnIndex].Value) < Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
                        {

                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.Red;
                        }


                    }
                }
            }
            catch { }

        }

        private void MyDataGridViewInitializationMethod()
        {
            this.dgvShipTracker.DataError += this.DataGridView_DataError;
            dgvShipTracker.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(dataGridView_EditingControlShowing);
        }

        private void dataGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
        }


        void DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
        }


        private void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)22)
            {
                string s = Clipboard.GetText();
                string[] lines = s.Split('\n');
                int row = dgvShipTracker.CurrentCell.RowIndex;
                int col = dgvShipTracker.CurrentCell.ColumnIndex;

                //   if ( dgvShipTracker[col, row].Value != "")
                //   dgvShipTracker[col, row].Value = "";
                // MessageBox.Show(dgvShipTracker[col, row].Value.ToString());

                try
                {
                    foreach (string line in lines)
                    {

                        if (row < dgvShipTracker.RowCount && line.Length > 0)
                        {
                            string[] cells = line.Split('\t');
                            for (int i = 0; i < cells.GetLength(0); ++i)
                            {
                                if (col + i < this.dgvShipTracker.ColumnCount)
                                {
                                    dgvShipTracker[col + i, row].Value = Convert.ChangeType(cells[i], dgvShipTracker[col + i, row].ValueType);
                                }
                                else
                                {
                                    break;
                                }
                            }
                            row++;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Column Data Type doesn't match");
                }
                // 
            }

        }

        private void dgvShipTracker_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                string s = Clipboard.GetText();
                string[] lines = s.Split('\n');
                int row = dgvShipTracker.CurrentCell.RowIndex;
                int col = dgvShipTracker.CurrentCell.ColumnIndex;
                try
                {
                    for (int rownum = 0; rownum < lines.Length - 1; rownum++)
                    {
                        dtShipTracker.Rows.Add(dtShipTracker.NewRow());
                        dgvShipTracker.DataSource = dtShipTracker;
                    }

                    foreach (string line in lines)
                    {
                        if (row < dgvShipTracker.RowCount && line.Length > 0)
                        {
                            string[] cells = line.Split('\t');
                            for (int i = 0; i < cells.GetLength(0); ++i)
                            {
                                if (col + i < this.dgvShipTracker.ColumnCount)
                                {
                                    dgvShipTracker[col + i, row].Value = Convert.ChangeType(cells[i], dgvShipTracker[col + i, row].ValueType);
                                }
                                else
                                {
                                    break;
                                }
                            }
                            row++;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Column data type doesn't match. Please check ");
                }
            }
        }

        DataTable tblCarrier;
        public void getCarrier()
        {
            DataSet ds = getDataSet("select * from vca_tracking_links ");
            tblCarrier = ds.Tables[0];
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            DataTable dtSave = dgvShipTracker.DataSource as DataTable;
            string sqlShipTracker = " select [GPN],[PACKING SLIP],[TRACKING NUM]  from VCA_SHIP_TRACKER_APP  ";
            dtdbShipTracker = getDataSet(sqlShipTracker).Tables[0];
            string strSave = "";

            DataRow[] ins = dtSave.Select(" id is  null or id=0");
            DataTable dtdup = new DataTable();
            dtdup = dtSave.Clone();
            //  ID,vendor as VENDOR,ship_id as [SHIP ID],received as [RECEIVED],eta_date as [ETA DATE],po_num as [PO NUM],gpn as [GPN],ship_to as [SHIP TO],ship_date as [SHIP DATE],ship_qty as [SHIP QTY], packing_slip as [PACKING SLIP], ship_method as [SHIP METHOD],carrier as [CARRIER], tracking_no as [TRACKING NUM] FROM         vca_ship_tracker_for_vcapp  where vendor<>''";
            bool duplicateRow = false;
            foreach (DataRow drins in ins)
            {
                DataRow[] dupRow = dtdbShipTracker.Select("[GPN]='" + drins["GPN"] + "' and [TRACKING NUM] like '" + drins["TRACKING NUM"].ToString().Trim().Replace("\r", "") + "' and  [PACKING SLIP]='" + drins["PACKING SLIP"] + "'");
                if (dupRow.Length > 0)
                {
                    DataRow drnew = dtdup.NewRow();
                    drnew["GPN"] = dupRow[0]["GPN"];

                    foreach (DataGridViewRow dr in dgvShipTracker.Rows)
                    {
                        if (dr.Cells["GPN"].Value != null)
                        {
                            string gridGPN = dr.Cells["GPN"].Value.ToString();
                            string dupRowGPN = dupRow[0]["GPN"].ToString();

                            string gridPackingSlip = dr.Cells["PACKING SLIP"].Value.ToString();
                            string dupRowPackingSlip = dupRow[0]["PACKING_SLIP"].ToString();

                            string gridTrackingNum = dr.Cells["TRACKING NUM"].Value.ToString().Replace("\r", "").Trim();
                            string dupRowTrackingNum = dupRow[0]["TRACKING_NO"].ToString().Replace("\r", "").Trim();

                            if (gridGPN == dupRowGPN)
                                if (gridPackingSlip == dupRowPackingSlip)
                                    if (gridTrackingNum == dupRowTrackingNum)
                                    {
                                        dgvShipTracker.Rows[dr.Index].Cells[0].Style.BackColor = Color.Red;
                                        dgvShipTracker.Rows[dr.Index].Cells[1].Style.ForeColor = Color.Red;
                                        dgvShipTracker.Rows[dr.Index].Cells[2].Style.ForeColor = Color.Red;
                                        dgvShipTracker.Rows[dr.Index].Cells[3].Style.ForeColor = Color.Red;
                                        dgvShipTracker.Rows[dr.Index].Cells[4].Style.ForeColor = Color.Red;
                                        dgvShipTracker.Rows[dr.Index].Cells[5].Style.ForeColor = Color.Red;
                                        dgvShipTracker.Rows[dr.Index].Cells[6].Style.ForeColor = Color.Red;
                                        dgvShipTracker.Rows[dr.Index].Cells[7].Style.ForeColor = Color.Red;
                                        dgvShipTracker.Rows[dr.Index].Cells[8].Style.ForeColor = Color.Red;
                                        dgvShipTracker.Rows[dr.Index].Cells[9].Style.ForeColor = Color.Red;
                                        duplicateRow = true;
                                    }
                        }
                    }
                }
            }
            if (duplicateRow == true)
            {
                MessageBox.Show("Please check the heightlighted row");
            }
            else
            {
                foreach (DataRow dr in dtSave.Rows)
                {
                    int id = 0;
                    string ship_id = "";

                    string received = "";

                    if (dr["ID"].ToString() != "")
                        id = Convert.ToInt32(dr["ID"].ToString());

                    if (dr["SHIP ID"].ToString() != "")
                        ship_id = dr["SHIP ID"].ToString();

                    if (dr["RECEIVED"].ToString() != "")
                        received = dr["RECEIVED"].ToString();


                    if (ship_id != "")
                        //if (id == 0)
                        //    strSave = strSave + " insert into vca_ship_tracker(ship_id,po_num,gpn,ship_date,ship_qty,ship_to,ship_method,eta_date,packing_slip,carrier,tracking_no,received) values('" + ship_id.ToString() + "','" + po_num.ToString() + "','" + gpn + "','" + ship_date + "','" + ship_qty + "','" + ship_to + "','" + ship_method + "','" + eta_date + "','" + packing_slip + "','" + carrier + "','" + tracking_no + "','" + received.Substring(0, 1) + "')";
                        //else
                        //    strSave = strSave + "  update vca_ship_tracker set ship_id='" + ship_id.ToString() + "',po_num='" + po_num + "',gpn='" + gpn + "',ship_qty='" + ship_qty + "',ship_to='" + ship_to + "',ship_method='" + ship_method + "',eta_date='" + eta_date + "',ship_date='" + ship_date + "',packing_slip='" + packing_slip + "',carrier='" + carrier + "',tracking_no='" + tracking_no + "',received='" + received + "'  where id=" + id;


                        strSave = strSave + "  update [VCA_ShipTrack_Table] set SHP_Recvd='" + received + "'  where AUTOID=" + id;



                }

                // MessageBox.Show("save" + strSave);
                SqlConnection cn = new SqlConnection(constr);
                try
                {
                    cn.Open();
                    SqlCommand cmd = new SqlCommand(strSave, cn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("ShipTracker has been updated successfully");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ShipTracker has been failed to save the data ");
                }
                finally
                {
                    cn.Dispose();
                    cn.Close();
                }
            }
        }

        private void dgvShipTracker_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


            if (e.RowIndex >= 0)
            {
                if (dgvShipTracker.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewLinkCell)
                {
                    if (dgvShipTracker.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "TRACK")
                    {
                        string Carrier = dgvShipTracker.Rows[e.RowIndex].Cells["CARRIER"].Value.ToString();
                        string sUrl = tblCarrier.Select("Carrier='" + Carrier.Replace("\r", "") + "'")[0]["link"].ToString();
                        string TrackNum = dgvShipTracker.Rows[e.RowIndex].Cells["TRACKING NUM"].Value.ToString();

                        sUrl = sUrl.Replace("~", TrackNum);
                        System.Diagnostics.Process.Start(sUrl);
                        //ProcessStartInfo sInfo = new ProcessStartInfo(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe");
                        //sInfo.UseShellExecute = false;
                        // sInfo.Arguments = sUrl;
                        // sInfo.RedirectStandardOutput = true;
                        //  Process.Start( sInfo);
                        // sInfo.Start(sUrl);
                    }
                    else
                    {
                        string id = dgvShipTracker.Rows[e.RowIndex].Cells[0].Value.ToString();
                        if (id != "" && id != null && id != "0")
                            deleteShipTracker(id);
                    }
                }


                if (dgvShipTracker.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewCheckBoxCell)
                {

                    DataGridViewCheckBoxCell chk = dgvShipTracker.Rows[e.RowIndex].Cells["Received_check"] as DataGridViewCheckBoxCell;
                    string val = chk.FormattedValue.ToString();
                    string val1 = chk.TrueValue.ToString();

                    if (dgvShipTracker.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected.ToString() == "Y")
                        dgvShipTracker.Rows[e.RowIndex].Cells["Received"].Value = "N";
                    else
                        dgvShipTracker.Rows[e.RowIndex].Cells["Received"].Value = "Y";

                }
            }

        }



        private void chkItems_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[3];
                if (chk.Value == chk.TrueValue)
                {
                    chk.Value = chk.FalseValue;
                }
                else
                {
                    chk.Value = chk.TrueValue;
                }
            }
        }

        public void deleteShipTracker(string id)
        {
            string sqlstr = "delete from vca_shiptrack_table where AutoId=" + id + "  select @@error ";
            DataSet ds = getDataSet(sqlstr);

            if (ds.Tables[0].Rows[0][0].ToString() == "0")
                MessageBox.Show("Record has been deleted successfully");
            else
                MessageBox.Show("Error: Failed to delete record");

            loadDatagrid();
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            //   getAllCommit(cmbVendorCommit.SelectedValue.ToString());
            Cursor.Current = Cursors.Default;
        }

        //public void getAllCommit(string vendor)
        //{
        //    dgvCommit.Columns.Clear();
        //    string sqlstr = "SELECT distinct vendor as Vendor,com_id  as [Commit ID], upload_date as [Upload Date], req_doc_date as [DOC Date]  FROM   vca_req_com where " +/*dbo.udf_GetISOWeekNumberFromDate(upload_date)=dbo.udf_GetISOWeekNumberFromDate(getdate()) and */" vendor='" + vendor + "' and req_doc_date>getdate()-30  order by req_doc_date desc ";
        //    DataSet ds=getDataSet(sqlstr);
        //    if (ds.Tables[0].Rows.Count > 0)
        //    {
        //        dgvCommit.DataSource = ds.Tables[0];
        //        DataGridViewLinkColumn btnDeleteCommit = new DataGridViewLinkColumn();
        //        btnDeleteCommit.UseColumnTextForLinkValue = true;
        //        btnDeleteCommit.Text = "Delete";
        //        // links.hr
        //        btnDeleteCommit.ActiveLinkColor = Color.White;
        //        btnDeleteCommit.LinkBehavior = LinkBehavior.SystemDefault;
        //        btnDeleteCommit.LinkColor = Color.Blue;
        //        btnDeleteCommit.TrackVisitedState = true;
        //        btnDeleteCommit.VisitedLinkColor = Color.YellowGreen;
        //        dgvCommit.Columns.Add(btnDeleteCommit);

        //    }

        //}

        private void dgvCommit_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex >= 0)
            //{
            //    if (dgvCommit.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewLinkCell)
            //    {
            //        if (dgvCommit.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "Delete")
            //        {
            //            string id = dgvCommit.Rows[e.RowIndex].Cells[1].Value.ToString();
            //            deleteCommit(id);
            //        }                   
            //    }
            //}
        }

        //    public void  deleteCommit(string id)
        //    {
        //           string sqlstrDeleteCommit = " delete from vca_req_com where com_id=" + id + "  select @@error ";
        //           DataSet ds=getDataSet(sqlstrDeleteCommit);
        //           if (ds.Tables[0].Rows[0][0].ToString() == "0")
        //               MessageBox.Show("Commit ID has been deleted");
        //           else
        //               MessageBox.Show("Faile to delete commit id");
        //           getAllCommit(cmbVendorCommit.SelectedValue.ToString());

        //}

        private void btnDownloadGrid_Click_1(object sender, EventArgs e)
        {
            if (lstWorksheet.SelectedItems.Count > 0)
            {

                string id = lstWorksheet.SelectedItems[0].SubItems[1].Text.ToString();
                filename = lstWorksheet.SelectedItems[0].SubItems[0].Text.ToString();
                Google.Apis.Drive.v3.Data.File file1 = service.Files.Get(id).Execute();
                saveTo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\currentVendorCommit.xlsx";
                saveTo = saveTo.Replace("file:\\", "");
                //  saveTo = "\\\\Mvfile\\mv reports\\Master Data & Forms\\PlanningExecutionForms\\DemandForecast\\"+lstWorksheet.SelectedItems[0].SubItems[0].Text.ToString()+".xlsx";
                if (downloadfile(service, id, saveTo))
                    openExcelInSpreadSheet(saveTo, this.fpSpread1);
                //  SetWorksheetListView();
                // showInGrid();
            }

        }

        public bool downloadfile(DriveService service, string fileId, string fileName)
        {

            var request = service.Files.Export(fileId, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            // var request = service.Files.Export(fileId, "application/application/x-vnd.oasis.opendocument.spreadsheet");

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


        private void lstWorksheet_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
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


        public void loadComboVendor()
        {

            //cmbVendorUpload.DataSource = dsVendor.Tables[0];
            //cmbVendorUpload.DisplayMember = "Vendor";
            //cmbVendorUpload.ValueMember = "vend_id";

        }


        public void loadVendorPartMaster()
        {
            if (cmbVendPart.SelectedValue.ToString() == "")
            {
                MessageBox.Show("Please Select Vendor", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
                string strVendPart = " SELECT     vend_part_id, vend_id, part_num AS [PART NUM], short_desc AS [PART DESC], std_ship_meth AS [SHIP METHOD],mfg_lt as [MFG LT], plan_type AS [PLAN TYPE], curr_po_at [CURR PO AT], curr_po_nl as  [CURR PO NL], std_box_qty as [STD BOX QTY], std_pallet_qty as [STD PALLET QTY], ord_multiple as [ORD MULTIPLE], factory_fob as [FACTORY FOB], CASE WHEN active='Y' THEN 'ACTIVE' ELSE 'INACTIVE' END AS ACTIVE,AltPartNum as [ATL_PART_NUM]  FROM   vc_vend_part_with_alt WHERE   vend_id = " + cmbVendPart.SelectedValue;
                if (cmbStatus.SelectedIndex != 0)
                    strVendPart = strVendPart + " and active='" + cmbStatus.SelectedValue + "'";
                strVendPart = strVendPart + "  Order by Part_num ";
                grdVendPart.DataSource = null;

                try
                {
                    DataSet dsVendPart = getdataSet(strVendPart);
                    grdVendPart.DataSource = dsVendPart.Tables[0];

                    grdVendPart.DisplayLayout.Bands[0].Columns[0].Hidden = true;
                    grdVendPart.DisplayLayout.Bands[0].Columns[1].Hidden = true;

                    grdVendPart.DisplayLayout.Bands[0].Columns[2].Width = 100;
                    grdVendPart.DisplayLayout.Bands[0].Columns[3].Width = 150;
                    grdVendPart.DisplayLayout.Bands[0].Columns[4].Width = 150;
                    grdVendPart.DisplayLayout.Bands[0].Columns[5].Width = 100;
                    grdVendPart.DisplayLayout.Bands[0].Columns[6].Width = 90;
                    grdVendPart.DisplayLayout.Bands[0].Columns[7].Width = 90;
                    grdVendPart.DisplayLayout.Bands[0].Columns[8].Width = 90;
                    grdVendPart.DisplayLayout.Bands[0].Columns[9].Width = 90;
                    grdVendPart.DisplayLayout.Bands[0].Columns[10].Width = 90;
                    grdVendPart.DisplayLayout.Bands[0].Columns[11].Width = 90;
                    grdVendPart.DisplayLayout.Bands[0].Columns[12].Width = 100;
                    grdVendPart.DisplayLayout.Bands[0].Columns[13].Width = 100;
                    grdVendPart.DisplayLayout.Bands[0].Columns[14].Width = 80;

                    grdVendPart.DisplayLayout.Bands[0].Columns[15].Hidden = true;
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error :" + e.Message.ToString());
                }
            }

        }


        private void grdVendPart_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
        {
            if (e.Row.Cells.Exists("DELETE"))
            {
                string path = "image\\delete2.ico";
                Image image = Bitmap.FromFile(path);
                e.Row.Cells["DELETE"].Value = image;
            }

            if (e.Row.Cells.Exists("EDIT"))
            {
                string path = "image\\edit.png";
                Image image = Bitmap.FromFile(path);
                e.Row.Cells["EDIT"].Value = image;
            }

            if (e.Row.Cells.Exists("[ALT PART NUM]"))
            {
                string path = "image\\add.png";
                Image image = Bitmap.FromFile(path);
                e.Row.Cells["ALT PART NUM"].Value = image;
            }
        }


        public void clearAllVendorPart()
        {
            this.lblPartVendID.Text = "";
            this.txtVendorPartNum.Text = "";
            this.txtPartDesc.Text = "";
            this.txtCurrPOAT.Text = "";
            this.txtCurrPONL.Text = "";
            // this.txtCurrPOSG.Text = "";
            this.txtOrdMultiPle.Value = 0;
            this.txtStdBoxQTY.Value = 0;
            this.txtStdPallQTY.Value = 0;
            this.cmbFOB.SelectedIndex = 0;
            this.cmbShipMethod.SelectedIndex = 0;
            this.cmbPlanType.SelectedIndex = 0;
            this.txtAltPartNum.Text = "";
            this.txtMfgLT.Value = 0;
            chkActive.Checked = true;

        }

        public void loadShipMethod()
        {
            string strsqlShipMthod = " SELECT distinct ship_meth as disp,ship_meth as value  FROM [MIMDIST].[dbo].[m_pe_maintian_lt]  ";
            DataSet dsVendorShipMethod = getdataSet(strsqlShipMthod);
            DataRow dr = dsVendorShipMethod.Tables[0].NewRow();
            dr["value"] = "0";
            dr["disp"] = "Select";
            dsVendorShipMethod.Tables[0].Rows.InsertAt(dr, 0);
            cmbShipMethod.DataSource = dsVendorShipMethod.Tables[0];
            cmbShipMethod.DisplayMember = "disp";
            cmbShipMethod.ValueMember = "value";
        }

        public void loadFOB()
        {
            string strsqlFOB = " SELECT distinct fob as disp,fob as value  FROM [MIMDIST].[dbo].[m_pe_maintian_lt]  ";
            DataSet dsVendorFOB = getdataSet(strsqlFOB);
            DataRow dr = dsVendorFOB.Tables[0].NewRow();
            dr["value"] = "0";
            dr["disp"] = "Select";
            dsVendorFOB.Tables[0].Rows.InsertAt(dr, 0);
            cmbFOB.DataSource = dsVendorFOB.Tables[0];
            cmbFOB.DisplayMember = "disp";
            cmbFOB.ValueMember = "value";
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
            dr["value"] = "GOOGLE";
            dt.Rows.Add(dr);

            DataRow dr1 = dt.NewRow();
            dr1["Disp"] = "MIM";
            dr1["value"] = "MIM";
            dt.Rows.Add(dr1);

            // dsVendorShipMethod.Tables[0].Rows.InsertAt(dr, 0);
            this.cmbPlanType.DataSource = dt;
            cmbPlanType.DisplayMember = "disp";
            cmbPlanType.ValueMember = "value";

        }


        DataTable dtVendorCode;
        private void vendorCodeVendorMater()
        {
            string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
            string strqry = " select vendor_code,address_name from dbo.apmaster where status_type = 5 and vend_class_code = 'VENDOR' AND vendor_code NOT LIKE 'DECOM%' AND flag_1099 = 1 ";
            DataSet ds = getDataSet(strqry);
            dtVendorCode = ds.Tables[0];
            this.cmbVendorCode.DataSource = ds.Tables[0];
            cmbVendorCode.ValueMember = "vendor_code";
            cmbVendorCode.DisplayMember = "vendor_code";
            DataTable dtVendType = new DataTable();
            dtVendType.Columns.Add(new DataColumn("VendType"));
            DataRow dr = dtVendType.NewRow();
            dr["VendType"] = "VCAPP";
            dtVendType.Rows.Add(dr);
            dr = dtVendType.NewRow();
            dr["VendType"] = "DISCRETE";
            dtVendType.Rows.Add(dr);
            comboVendType.DataSource = dtVendType;
        }

        private void btnLoadVendor_Click(object sender, EventArgs e)
        {
            loadVendorMasterGrid();
            vendorCodeVendorMater();
        }

        public void loadVendorMasterGrid()
        {
            string sqlGrid = " select vend_id,ep_vend_code as [VENDOR CODE],vend_type as [VENDOR TYPE],shared_doc_name as [SHARED DOC NAME],[Goo_Integrated] as [GOO INTEGRATED] from vc_vend_master where active='Y' order by  ep_vend_code";
            grdVendor.DataSource = null;
            try
            {
                grdVendor.DataSource = getdataSet(sqlGrid).Tables[0];
                grdVendor.DisplayLayout.Bands[0].Columns[0].Hidden = true;
                grdVendor.DisplayLayout.Bands[0].Columns[3].Width = 250;
            }
            catch
            {

            }

        }

        public void loadshareddoc()
        {
            string sqlGrid = " select Distinct shared_doc_name as [SHARED DOC NAME] from vc_vend_master where active='Y' ";
            grdVendor.DataSource = null;
            try
            {
                grdVendor.DataSource = getdataSet(sqlGrid).Tables[0];
                grdVendor.DisplayLayout.Bands[0].Columns[0].Hidden = true;
                grdVendor.DisplayLayout.Bands[0].Columns[2].Width = 250;
            }
            catch
            {

            }

        }

        private void btnClearVendor_Click(object sender, EventArgs e)
        {
            clearAllVendor();
        }

        private void btnClearVendPart_Click(object sender, EventArgs e)
        {
            this.lblPartVendID.Text = "";
            this.txtVendorPartNum.Text = "";
            this.txtPartDesc.Text = "";
            this.txtCurrPOAT.Text = "";
            this.txtCurrPONL.Text = "";
            //  this.txtCurrPOSG.Text = "";
            this.txtOrdMultiPle.Value = 0;
            this.txtStdBoxQTY.Value = 0;
            this.txtStdPallQTY.Value = 0;
            this.cmbFOB.SelectedIndex = 0;
            this.cmbShipMethod.SelectedIndex = 0;
            this.cmbPlanType.SelectedIndex = 0;
        }

        private void grdVendor_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
            Infragistics.Win.UltraWinGrid.UltraGridLayout layout = e.Layout;
            Infragistics.Win.UltraWinGrid.UltraGridBand band = layout.Bands[0];

            Infragistics.Win.UltraWinGrid.UltraGridColumn editColumn = band.Columns.Add("EDIT");
            editColumn.DataType = typeof(Image);
            editColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Image;

            Infragistics.Win.UltraWinGrid.UltraGridColumn imageColumn = band.Columns.Add("DELETE");
            imageColumn.DataType = typeof(Image);
            imageColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Image;
        }

        private void grdVendor_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
        {
            if (e.Row.Cells.Exists("DELETE"))
            {
                string path = "image\\delete2.ico";
                Image image = Bitmap.FromFile(path);
                e.Row.Cells["DELETE"].Value = image;
            }

            if (e.Row.Cells.Exists("EDIT"))
            {
                string path = "image\\edit.png";
                Image image = Bitmap.FromFile(path);
                e.Row.Cells["EDIT"].Value = image;
            }


        }

        private void grdVendor_ClickCell(object sender, Infragistics.Win.UltraWinGrid.ClickCellEventArgs e)
        {
            if (e.Cell.Column.Index == 6)
            {
                if (e.Cell.Row.Cells[0].Value.ToString() != "")
                {
                    DeleteRowDataBase(e.Cell.Row.Cells[0].Value.ToString());
                }

                this.grdVendor.Rows[e.Cell.Row.Index].Delete();
            }

            if (e.Cell.Column.Index == 5)
            {
                lblAutoID.Text = grdVendor.Rows[e.Cell.Row.Index].Cells[0].Value.ToString();
                cmbVendorCode.Text = grdVendor.Rows[e.Cell.Row.Index].Cells[1].Value.ToString();
                txtVendorType.Text = grdVendor.Rows[e.Cell.Row.Index].Cells[3].Value.ToString();
                comboVendType.Text = grdVendor.Rows[e.Cell.Row.Index].Cells[2].Value.ToString();
                if (grdVendor.Rows[e.Cell.Row.Index].Cells[4].Value.ToString() == "Y")
                    this.checkBox1.Checked = true;
                else
                    this.checkBox1.Checked = false;

            }
        }

        public void DeleteRowDataBase(string Auto_id)
        {
            string sqlDelete = "update [MIMDIST].[dbo].[vc_vend_master] set active='N'  where vend_id=" + Auto_id.ToString();
            getDataSet(sqlDelete);

        }

        private void btnUpdateVendor_Click(object sender, EventArgs e)
        {
            string sqlStr = "";
            if (dtVendorCode.Select("vendor_code='" + cmbVendorCode.Text + "'").Length == 0)
            {
                MessageBox.Show("Vendor Code is not valid. Please check", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (txtVendorType.Text == "")
            {
                MessageBox.Show("Shared doc name can not be blank", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                string goo_integrated = "N";
                if (checkBox1.Checked)
                    goo_integrated = "Y";
                sqlStr = " if exists ( select 1 from vc_vend_master where ep_vend_code='" + cmbVendorCode.Text.Trim() + "'  ) update vc_vend_master set goo_integrated='" + goo_integrated + "', ep_vend_code='" + cmbVendorCode.Text + "',vend_type='" + this.comboVendType.Text.Trim() + "',shared_doc_name='" + txtVendorType.Text + "' where ep_vend_code='" + cmbVendorCode.Text.Trim() + "' else insert into vc_vend_master(ep_vend_code,shared_doc_name,vend_type,goo_integrated) values('" + cmbVendorCode.Text + "','" + txtVendorType.Text + "','" + this.comboVendType.Text + "','" + goo_integrated + "') ";
                try
                {
                    SqlConnection cn = new SqlConnection(constr);
                    cn.Open();
                    SqlCommand cmd = new SqlCommand(sqlStr, cn);
                    cmd.ExecuteNonQuery();
                    cn.Close();
                    MessageBox.Show("Vendor has  been updated Successfully");
                    clearAllVendor();
                    loadVendorMasterGrid();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Vendor has not been deleted due to some error. Please check", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        private void clearAllVendor()
        {
            cmbVendorCode.Text = "";
            txtVendorType.Text = "";
            this.comboVendType.Text = "";
        }

        private void grdVendPart_ClickCell(object sender, Infragistics.Win.UltraWinGrid.ClickCellEventArgs e)
        {

            if (e.Cell.Column.Index == 3)
            {
                AlternatePartPopup pop = new AlternatePartPopup(Convert.ToInt32(grdVendPart.Rows[e.Cell.Row.Index].Cells["vend_part_id"].Value.ToString()));
                pop.ShowDialog();
            }
            else if (e.Cell.Column.Index == 16)
            {
                this.lblPartVendID.Text = grdVendPart.Rows[e.Cell.Row.Index].Cells["vend_part_id"].Value.ToString();
                txtVendorPartNum.Text = grdVendPart.Rows[e.Cell.Row.Index].Cells["PART NUM"].Value.ToString();
                txtAltPartNum.Text = grdVendPart.Rows[e.Cell.Row.Index].Cells["ATL_PART_NUM"].Value.ToString();
                txtPartDesc.Text = grdVendPart.Rows[e.Cell.Row.Index].Cells["PART DESC"].Value.ToString();
                this.cmbShipMethod.Text = grdVendPart.Rows[e.Cell.Row.Index].Cells["SHIP METHOD"].Value.ToString();
                this.txtMfgLT.Text = grdVendPart.Rows[e.Cell.Row.Index].Cells["MFG LT"].Value.ToString();

                cmbPlanType.Text = grdVendPart.Rows[e.Cell.Row.Index].Cells["PLAN TYPE"].Value.ToString();
                txtCurrPOAT.Text = grdVendPart.Rows[e.Cell.Row.Index].Cells["CURR PO AT"].Value.ToString();
                txtCurrPONL.Text = grdVendPart.Rows[e.Cell.Row.Index].Cells["CURR PO NL"].Value.ToString();
                //   txtCurrPOSG.Text = grdVendPart.Rows[e.Cell.Row.Index].Cells["CURR PO SG"].Value.ToString();

                this.txtStdBoxQTY.Value = grdVendPart.Rows[e.Cell.Row.Index].Cells["STD BOX QTY"].Value.ToString();
                this.txtStdPallQTY.Value = grdVendPart.Rows[e.Cell.Row.Index].Cells["STD PALLET QTY"].Value.ToString();
                this.txtOrdMultiPle.Value = grdVendPart.Rows[e.Cell.Row.Index].Cells["ORD MULTIPLE"].Value.ToString();

                if (this.grdVendPart.Rows[e.Cell.Row.Index].Cells["FACTORY FOB"].Value.ToString() == "ASIA")
                    this.cmbFOB.SelectedIndex = 1;
                else if (this.grdVendPart.Rows[e.Cell.Row.Index].Cells["FACTORY FOB"].Value.ToString() == "EUROPE")
                    this.cmbFOB.SelectedIndex = 2;
                else if (this.grdVendPart.Rows[e.Cell.Row.Index].Cells["FACTORY FOB"].Value.ToString() == "US")
                    this.cmbFOB.SelectedIndex = 3;

                if (this.grdVendPart.Rows[e.Cell.Row.Index].Cells["PLAN TYPE"].Value.ToString() == "G")
                    cmbPlanType.Text = "GOOGLE";
                if (this.grdVendPart.Rows[e.Cell.Row.Index].Cells["PLAN TYPE"].Value.ToString() == "M")
                    cmbPlanType.Text = "MIM";

                if (this.grdVendPart.Rows[e.Cell.Row.Index].Cells["ACTIVE"].Value.ToString() == "ACTIVE")
                    chkActive.Checked = true;
                else
                    chkActive.Checked = false;

            }
        }

        private void grdVendPart_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
            Infragistics.Win.UltraWinGrid.UltraGridLayout layout = e.Layout;
            Infragistics.Win.UltraWinGrid.UltraGridBand band = layout.Bands[0];

            Infragistics.Win.UltraWinGrid.UltraGridColumn editColumn = band.Columns.Add("EDIT");
            editColumn.DataType = typeof(Image);
            editColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Image;

            Infragistics.Win.UltraWinGrid.UltraGridColumn AltPartNumColumn = band.Columns.Insert(3, "A PART #");
            AltPartNumColumn.DataType = typeof(Image);
            AltPartNumColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Image;
            AltPartNumColumn.Width = 10;

        }

        private void grdVendPart_InitializeRow_1(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
        {
            if (e.Row.Cells.Exists("EDIT"))
            {
                string path = "image\\edit.png";
                Image image = Bitmap.FromFile(path);
                e.Row.Cells["EDIT"].Value = image;
            }
            if (e.Row.Cells.Exists("A PART #"))
            {
                string path = "image\\add.png";
                Image image = Bitmap.FromFile(path);
                e.Row.Cells["A PART #"].Value = image;
            }
        }

        private void btnUpdateVendorPart_Click(object sender, EventArgs e)
        {
            string sqlStr = "";
            string act = "N";
            if (chkActive.Checked == true)
                act = "Y";

            if (!this.validatePart(this.txtVendorPartNum.Text.Trim()))
            {
                MessageBox.Show("Part number doesn't exist in epicor", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.txtVendorPartNum.SelectAll();
                this.txtVendorPartNum.Focus();
            }
            else if (cmbPlanType.SelectedIndex == 0)
            {
                MessageBox.Show("Please select the Plan Type", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbPlanType.SelectAll();
                cmbPlanType.Focus();
            }
            else if (cmbFOB.SelectedIndex == 0)
            {
                MessageBox.Show("Please select the FOB", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbFOB.SelectAll();
                cmbFOB.Focus();
            }
            else if (cmbShipMethod.SelectedIndex == 0)
            {
                MessageBox.Show("Please select the ship method", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbShipMethod.SelectAll();
                cmbShipMethod.Focus();
            }
            else
            {
                if (this.lblPartVendID.Text == "" || this.lblPartVendID.Text == "0")
                {
                    sqlStr = "insert into vc_vend_part(vend_id, part_num, alt_part_num,short_desc, std_ship_meth,mfg_lt,plan_type, curr_po_at , curr_po_nl,  std_box_qty, std_pallet_qty, ord_multiple, factory_fob) values('" + this.cmbVendPart.SelectedValue + "','" + this.txtVendorPartNum.Text.Trim() + "','" + this.txtAltPartNum.Text.Trim() + "','" + this.txtPartDesc.Text.Trim() + "','" + this.cmbShipMethod.Text.Trim() + "','" + this.txtMfgLT.Value.ToString().Trim() + "','" + this.cmbPlanType.SelectedValue.ToString().Trim() + "','" + this.txtCurrPOAT.Text.Trim() + "','" + this.txtCurrPONL.Text.Trim() + "','" + this.txtStdBoxQTY.Value.ToString().Trim() + "','" + this.txtStdPallQTY.Value.ToString().Trim() + "','" + this.txtOrdMultiPle.Value.ToString().Trim() + "','" + this.cmbFOB.Text.Trim() + "')";

                }
                else
                {
                    sqlStr = "update vc_vend_part set part_num='" + this.txtVendorPartNum.Text + "',alt_part_num='" + this.txtAltPartNum.Text.Trim() + "',mfg_lt='" + this.txtMfgLT.Value.ToString().Trim() + "',short_desc='" + txtPartDesc.Text.Trim() + "', std_ship_meth='" + cmbShipMethod.Text.Trim() + "',curr_po_at='" + txtCurrPOAT.Text.Trim() + "', curr_po_nl='" + txtCurrPONL.Text.Trim() + "',plan_type='" + this.cmbPlanType.SelectedValue.ToString().Trim() + "',Active='" + act + "',std_box_qty='" + this.txtStdBoxQTY.Value.ToString().Trim() + "', std_pallet_qty='" + this.txtStdPallQTY.Value.ToString().Trim() + "', ord_multiple='" + this.txtOrdMultiPle.Value.ToString().Trim() + "', factory_fob='" + this.cmbFOB.Text.Trim() + "'  where     vend_part_id=" + lblPartVendID.Text.Trim();
                }

                try
                {
                    SqlConnection cn = new SqlConnection(constr);
                    cn.Open();
                    SqlCommand cmd = new SqlCommand(sqlStr, cn);
                    cmd.ExecuteNonQuery();
                    cn.Close();

                    clearAllVendorPart();
                    loadVendorPartMaster();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Part has not been inserted due to some error. Please check", "Saving Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
        }

        private void btnPipe_Click(object sender, EventArgs e)
        {
            txtAltPartNum.Text = txtAltPartNum.Text + "|";
            txtAltPartNum.Focus();
            txtAltPartNum.SelectionStart = txtAltPartNum.Text.Length; // add some logic if length is 0
            txtAltPartNum.SelectionLength = 0;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void btnSaveCommit_Click(object sender, EventArgs e)
        {
            string vendor = fpSpread1.Sheets["Forecast"].Cells[1, 1].Value.ToString();
            string req_id = fpSpread1.Sheets["Forecast"].Cells[2, 1].Value.ToString();
            int part_start_row = 4;
            int part_start_col = 1;
            int part_num_row = 12;
            int AT_row = 8;
            int week_num_col = 5;
            int rowcount = fpSpread1.Sheets["Forecast"].Rows.Count;
            string sqlQuery = "";
            bool error = false;
            string errorPartNum = "";
            string errorWeekNum = "";
            string errorSite = "";
            while (part_start_row < rowcount)
            {
                string part_num = "";
                if (fpSpread1.Sheets["Forecast"].Cells[part_start_row, 1].Value != null)
                    part_num = fpSpread1.Sheets["Forecast"].Cells[part_start_row, 1].Value.ToString();

                if (part_num == "")
                    goto Complete;
                string week_num = "";

                for (int cols = week_num_col; cols < 41; cols++)
                {
                    if (fpSpread1.Sheets["Forecast"].Cells[1, cols].Value.ToString().Length < 7)
                    {
                        string front = fpSpread1.Sheets["Forecast"].Cells[1, cols].Value.ToString().Substring(0, 5);
                        string end = fpSpread1.Sheets["Forecast"].Cells[1, cols].Value.ToString().Substring(5);
                        week_num = front + "0" + end;
                    }
                    else
                        week_num = fpSpread1.Sheets["Forecast"].Cells[1, cols].Value.ToString();
                    int com_qty_at = 0;
                    int com_qty_nl = 0;
                    int com_qty_sg = 0;
                    if ((fpSpread1.Sheets["Forecast"].Cells[AT_row, cols].Value != null) && (fpSpread1.Sheets["Forecast"].Cells[AT_row, cols].Value.ToString().Trim() != ""))
                        try
                        {
                            com_qty_at = Convert.ToInt32(fpSpread1.Sheets["Forecast"].Cells[AT_row, cols].Value.ToString());
                        }
                        catch
                        {
                            fpSpread1.Sheets["Forecast"].Cells[AT_row, cols].BackColor = Color.Red;
                            errorPartNum = part_num;
                            errorWeekNum = week_num;
                            errorSite = "AT";
                            break;

                        }
                    if ((fpSpread1.Sheets["Forecast"].Cells[AT_row + 1, cols].Value != null) && (fpSpread1.Sheets["Forecast"].Cells[AT_row + 1, cols].Value.ToString().Trim() != ""))
                        try
                        {
                            com_qty_nl = Convert.ToInt32(fpSpread1.Sheets["Forecast"].Cells[AT_row + 1, cols].Value.ToString());
                        }
                        catch
                        {
                            fpSpread1.Sheets["Forecast"].Cells[AT_row + 1, cols].BackColor = Color.Red;
                            errorPartNum = part_num;
                            errorWeekNum = week_num;
                            errorSite = "NL";
                            error = true;
                            break;
                        }
                    if ((fpSpread1.Sheets["Forecast"].Cells[AT_row + 2, cols].Value != null) && (fpSpread1.Sheets["Forecast"].Cells[AT_row + 2, cols].Value.ToString().Trim() != ""))

                        try
                        {
                            com_qty_sg = Convert.ToInt32(fpSpread1.Sheets["Forecast"].Cells[AT_row + 2, cols].Value.ToString());
                        }
                        catch
                        {
                            fpSpread1.Sheets["Forecast"].Cells[AT_row + 2, cols].BackColor = Color.Red;
                            errorPartNum = part_num;
                            errorWeekNum = week_num;
                            errorSite = "SG";
                            error = true;
                            break;
                        }
                    sqlQuery = sqlQuery + "   update  [dbo].[VCA_ReqCom_Table] set req_commit_qty= " + com_qty_at + ",Req_commitRecvFlag='Y', Req_commit_date=getdate()  where REQ_ID=" + req_id + "  and REQ_VendCode='" + vendor + "' and  REQ_GPN='" + part_num + "' and REQ_SiteCode='AT' and REQ_YRWK='" + week_num + "'";
                    sqlQuery = sqlQuery + "   update  [dbo].[VCA_ReqCom_Table] set req_commit_qty= " + com_qty_nl + ",Req_commitRecvFlag='Y' , Req_commit_date=getdate()  where REQ_ID=" + req_id + "  and REQ_VendCode='" + vendor + "' and  REQ_GPN='" + part_num + "' and REQ_SiteCode='NL' and REQ_YRWK='" + week_num + "'";
                    sqlQuery = sqlQuery + "   update  [dbo].[VCA_ReqCom_Table] set req_commit_qty= " + com_qty_sg + ",Req_commitRecvFlag='Y' , Req_commit_date=getdate()  where REQ_ID=" + req_id + "  and REQ_VendCode='" + vendor + "' and  REQ_GPN='" + part_num + "' and REQ_SiteCode='SG' and REQ_YRWK='" + week_num + "'";
                }

                part_start_row = part_start_row + part_num_row;
                AT_row = AT_row + part_num_row;
            }



        Complete:
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlTransaction tran = cn.BeginTransaction();
            sqlQuery = sqlQuery + " update  [dbo].[VCA_ReqCom_Table] set Req_commitRecvFlag='Y' , Req_commit_date=getdate()  where REQ_ID=" + req_id;

            SqlCommand cmd = new SqlCommand(sqlQuery, cn);
            cmd.CommandTimeout = 0;
            if (!error)
            {
                try
                {
                    cmd.Transaction = tran;
                    cmd.ExecuteNonQuery();
                    tran.Commit();
                    MessageBox.Show("Commit has been saved successfully", "SUCCESS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception exp)
                {
                    tran.Rollback();
                    string Error = exp.Message.ToString();
                    MessageBox.Show("Commit save failed ", "Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                MessageBox.Show("There is an error in data. Look for the cell with red background! Part Num=" + errorPartNum + "  Week# =" + errorWeekNum + " Site =" + errorSite, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            cmd.Dispose();
            cn.Close();

        }


        string vendor = "";
        private void button1_Click_2(object sender, EventArgs e)
        {
            fpSpreadUploadSharedoc.Open("\\\\mvfile\\MV REPORTS\\Master Data & Forms\\PlanningExecutionForms\\VCAPP\\VendorRequstTemplate.xml");
            if (txtReqID.Value.ToString() == "0")
                getRequestId();
            if (txtReqID.Value.ToString() != "0" && txtReqID.Value.ToString() != "")
            {
                string sqlGetRequest = "   select * from  [dbo].[VCA_Demand_forecast_data_app] where REQ_ID =" + txtReqID.Value.ToString() + " order by ReQ_GPN,REQ_SiteCode, Req_YrWK ";
                DataSet ds = getDataSet(sqlGetRequest);
                if (ds.Tables.Count > 0)
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        vendor = ds.Tables[0].Rows[0]["REQ_VendCode"].ToString();
                        fpSpreadUploadSharedoc.Sheets[0].Cells[1, 1].Value = ds.Tables[0].Rows[0]["REQ_VendCode"].ToString();
                        fpSpreadUploadSharedoc.Sheets[0].Cells[2, 1].Value = ds.Tables[0].Rows[0]["REQ_ID"].ToString();
                        fpSpreadUploadSharedoc.Sheets[0].Cells[0, 1].Value = Convert.ToDateTime(ds.Tables[0].Rows[0]["REQ_createdate"].ToString());

                        FarPoint.Win.Spread.CellType.DateTimeCellType datecell = new FarPoint.Win.Spread.CellType.DateTimeCellType();
                        datecell.DateTimeFormat = FarPoint.Win.Spread.CellType.DateTimeFormat.ShortDate;
                        FarPoint.Win.Spread.CellType.NumberCellType numcell = new FarPoint.Win.Spread.CellType.NumberCellType();
                        numcell.Separator = ",";
                        fpSpreadUploadSharedoc.Sheets[0].Cells[0, 1].CellType = datecell;
                        int part_start_row = 4;
                        int part_start_col = 1;
                        int part_num_row = 12;
                        int AT_row = 8;
                        int week_num_col = 5;
                        int numCol = 60;
                        int rowcount = fpSpreadUploadSharedoc.Sheets["Forecast"].Rows.Count;
                        string sqlQuery = "";
                        int part_count = 0;
                        DataView view = new DataView(ds.Tables[0]);
                        DataTable distinctParts = view.ToTable(true, "REQ_GPN");

                        foreach (DataRow drPart in distinctParts.Rows)
                        {
                            week_num_col = 5;
                            int week_num_col_at = week_num_col;
                            int week_num_col_nl = week_num_col;
                            int week_num_col_sg = week_num_col;

                            part_count = part_count + 1;
                            fpSpreadUploadSharedoc.Sheets[0].Cells[part_start_row, 1].Value = drPart[0].ToString();
                            if (part_count < distinctParts.Rows.Count)
                                fpSpreadUploadSharedoc.Sheets[0].CopyRange(part_start_row, 0, part_start_row + part_num_row, 0, part_num_row, numCol, false);
                            int at_row = 0;
                            int nl_row = 0;
                            int sg_row = 0;
                            foreach (DataRow drWeekReq in ds.Tables[0].Select("REQ_GPN='" + drPart[0].ToString() + "'"))
                            {
                                fpSpreadUploadSharedoc.Sheets[0].Cells[part_start_row + 2, 1].Value = drWeekReq["alt_part_num"].ToString();
                                fpSpreadUploadSharedoc.Sheets[0].Cells[part_start_row + 1, 1].Value = drWeekReq["short_desc"].ToString();
                                fpSpreadUploadSharedoc.Sheets[0].Cells[part_start_row + 9, 1].Value = drWeekReq["curr_po_nl"].ToString();
                                fpSpreadUploadSharedoc.Sheets[0].Cells[part_start_row + 8, 1].Value = drWeekReq["curr_po_at"].ToString();
                                //   fpSpreadUploadSharedoc.Sheets[0].Cells[part_start_row + 10, 1].Value = drWeekReq["curr_po_sg"].ToString();
                                if (drWeekReq["REQ_Sitecode"].ToString() == "AT")
                                {
                                    if (at_row != 0)
                                    {
                                        string qtr = drWeekReq["Quarter"].ToString();
                                        string wk_num = drWeekReq["Req_YRWK"].ToString();
                                        DateTime start_date = Convert.ToDateTime(drWeekReq["Start_date"].ToString());
                                        fpSpreadUploadSharedoc.Sheets["Forecast"].Cells[0, week_num_col].Value = qtr;
                                        fpSpreadUploadSharedoc.Sheets["Forecast"].Cells[1, week_num_col].Value = wk_num;
                                        fpSpreadUploadSharedoc.Sheets["Forecast"].Cells[2, week_num_col].Value = start_date;
                                        //  fpSpreadUploadSharedoc.Sheets["Forecast"].Cells[2, week_num_col].CellType = datecell;
                                        part_start_row = part_start_row;
                                        //   fpSpreadUploadSharedoc.Sheets["Forecast"].Cells[part_start_row, week_num_col_at].CellType = numcell;
                                        string WkreqAT = "";
                                        WkreqAT = drWeekReq["Req_QTY"].ToString();
                                        if (WkreqAT != "0")
                                            fpSpreadUploadSharedoc.Sheets["Forecast"].Cells[part_start_row, week_num_col_at].Value = Convert.ToInt32(drWeekReq["Req_QTY"].ToString()).ToString("#,###,###");
                                        else
                                            fpSpreadUploadSharedoc.Sheets["Forecast"].Cells[part_start_row, week_num_col_at].Value = 0;

                                        week_num_col_at = week_num_col_at + 1;
                                        week_num_col = week_num_col + 1;
                                    }
                                    at_row = at_row + 1;
                                }
                                else if (drWeekReq["REQ_Sitecode"].ToString() == "NL")
                                {
                                    fpSpreadUploadSharedoc.Sheets["Forecast"].Cells[part_start_row + 1, week_num_col_nl].Value = 0;
                                    if (nl_row != 0)
                                    {
                                        string WkreqNL = "";
                                        WkreqNL = drWeekReq["Req_QTY"].ToString();
                                        if (WkreqNL != "0")
                                            fpSpreadUploadSharedoc.Sheets["Forecast"].Cells[part_start_row + 1, week_num_col_nl].Value = Convert.ToInt32(drWeekReq["Req_QTY"].ToString()).ToString("#,###,###");
                                        week_num_col_nl = week_num_col_nl + 1;
                                    }
                                    nl_row = nl_row + 1;
                                }
                                /*     else if (drWeekReq["REQ_Sitecode"].ToString() == "SG")
                                     {
                                         fpSpreadUploadSharedoc.Sheets["Forecast"].Cells[part_start_row + 2, week_num_col_sg].Value = 0;
                                         if (sg_row != 0)
                                         {
                                             string WkreqSG = "";
                                             WkreqSG = drWeekReq["Req_QTY"].ToString();
                                             if (WkreqSG != "0")
                                                 fpSpreadUploadSharedoc.Sheets["Forecast"].Cells[part_start_row + 2, week_num_col_sg].Value = Convert.ToInt32(drWeekReq["Req_QTY"].ToString()).ToString("#,###,###");
                                             week_num_col_sg = week_num_col_sg + 1;
                                         }
                                         sg_row = sg_row + 1;
                                     }*/
                            }
                            part_start_row = part_start_row + part_num_row;
                        }
                    }
            }
            else
                MessageBox.Show("Please enter Request ID or select vendor,week nummber and year ", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);


        }

        private void button4_Click(object sender, EventArgs e)
        {
            string filename = "Demand Forecast " + vendor + " " + this.cmbYearRequest.Text.ToString() + "W" + cmbWeekRequest.Text.ToString() + ".xlsx";
            fpSpreadUploadSharedoc.SaveExcel(filename);
            uploadfile(filename);
        }

        string VendorNamefolder = "";
        public string getFolderName(string vendor)
        {
            string sqlStr = " SELECT [vend_id]  ,[ep_vend_code]  ,[vend_type]  ,[shared_doc_name]  ,[Goo_Integrated]  FROM [MIMDIST].[dbo].[vc_vend_master] where active='Y' and ep_vend_code='" + vendor + "'";
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlStr, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet dsFolder = new DataSet();
            da.Fill(dsFolder);
            if (dsFolder.Tables.Count > 0)
                if (dsFolder.Tables[0].Rows.Count > 0)
                    return dsFolder.Tables[0].Rows[0]["shared_doc_name"].ToString();

            cn.Close();
            return "";

        }



        /*    public void moveALLchildToArchive(DriveService service, String folderId, String ArchivefolderId, string filestart,string docname)
            {
                this.lstWorksheet.Items.Clear();
                ChildrenResource.ListRequest request = service.Children.List(folderId);
                do
                {
                    try
                    {
                        request.Q = " title  contains '" + filestart + "'";
                        ChildList children = request.Execute();
                        ParentReference newParent = new ParentReference();
                        ParentReference oldParent = new ParentReference(); 

                        foreach (ChildReference child in children.Items)
                        {
                            Google.Apis.Drive.v2.Data.File file1 = service.Files.Get(child.Id).Execute();
                            if (file1.Title.Contains(filestart) && file1.Title.ToString() != docname)
                              {                                               
                                googleshareDoc = file1.Title;
                                newParent.Id = ArchivefolderId;                        
                                service.Parents.Insert(newParent, child.Id).Execute();
                                service.Parents.Delete(child.Id, folderId).Execute();                           
                              }
                            // }
                        }

                        request.PageToken = children.NextPageToken;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("An error occurred: " + e.Message);
                        request.PageToken = null;
                    }
                } while (!String.IsNullOrEmpty(request.PageToken));
        
            }*/
        public void uploadfile(string docname)
        {

            docname = docname.Replace(".xlsx", "");
            string folderName = getFolderName(vendor);
            string ArchivefolderName = vendor + " Archive";
            string Parentid = "";
            string ArchiveParentid = "";


            Parentid = getGIDByName(folderName);
            ArchiveParentid = getGIDByName(ArchivefolderName);
            string pageToken = "";
            fileLink = "";
            do
            {

                //  string parent_id = getGIDByName(vendor);
                var result = getFilesByParentId(Parentid, pageToken);
                foreach (var file in result.Files)
                {

                    if (file.Name.StartsWith("Demand Forecast"))
                    {
                        fileLink = file.Id;
                        break;
                    }
                }
                pageToken = result.NextPageToken;
            } while (pageToken != null);

            //  moveALLchildToArchive(service, Parentid, ArchiveParentid, "Demand Forecast", docname);
            //  fileLink= getFirstchildInFolder(service, Parentid, docname);
            if (fileLink != "")
                moveFile(fileLink, ArchiveParentid);
            try
            {
                Google.Apis.Drive.v3.Data.File body = new Google.Apis.Drive.v3.Data.File();
                body.Name = docname;
                body.Description = docname;
                body.MimeType = "application/vnd.google-apps.spreadsheet";
                body.Parents = new List<string> { Parentid };

                filename = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\" + docname + ".xlsx";// "E:\\Ruchi\\b2bdevMaintainGoogleDoc\\MaintainGoogleDoc\\MaintainGoogleDoc\\bin\\Debug\\MIM_Menlo Demand TrackerTemplate.xlsx";
                filename = filename.Replace("file:\\", " ");
                // MessageBox.Show("filename=" + filename);
                System.IO.Stream fileStream = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                byte[] byteArray = new byte[fileStream.Length];
                fileStream.Read(byteArray, 0, (int)fileStream.Length);
                System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
                //   if (fileLink == "")
                // {
                FilesResource.CreateMediaUpload request = service.Files.Create(body, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
                // request.Convert = true;
                request.Upload();
                //}
                // else
                // {
                //   FilesResource.UpdateMediaUpload request = service.Files.Update(body, fileLink, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
                // request.Convert = true;
                // request.Upload();
                // }

                MessageBox.Show("File has been uploaded in shared doc", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                fileStream.Close();
                fileStream.Dispose();
                updateUploadFlag();
            }
            catch (Exception exp)
            {
                MessageBox.Show("Exception to upload file", "Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {

            }

        }

        public void moveFile(string fileId, string toFolder)
        {

            // Retrieve the existing parents to remove
            var getRequest = service.Files.Get(fileId);
            getRequest.Fields = "parents";
            var file = getRequest.Execute();
            var previousParents = String.Join(",", file.Parents);
            // var getRequest1 = service.Files.Get(fileId);
            // var CopyRequest = service.Files.Copy(new Google.Apis.Drive.v3.Data.File(), fileId);
            //var Copyfile = CopyRequest.Execute();

            var updateRequest = service.Files.Update(new Google.Apis.Drive.v3.Data.File(), fileId);
            updateRequest.Fields = "id, parents";
            updateRequest.AddParents = toFolder;
            updateRequest.RemoveParents = previousParents;
            file = updateRequest.Execute();

        }


        /*   private void moveFileToArchive(DriveService driveService, string fileId, string folderId)
           {
                 File file =  driveService.Files.Get(fileId).Execute();
                 File targetFolder = driveService.Files.Get(folderId).Execute();   
                 file.Parents = new List<ParentReference>() { new ParentReference() { Id = folderId }};
                 File updatedFile = driveService.Files.Update(file, fileId).Execute();
           }*/

        private void btnExport_Click(object sender, EventArgs e)
        {
            string filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + " \\Demand Forecast " + vendor + ".xls";
            fpSpreadUploadSharedoc.SaveExcel(filename);
            Process.Start(filename);
        }

        private void txtReqID_ValueChanged(object sender, EventArgs e)
        {

        }
        string fname = "";
        private void btnShareDoc_Click(object sender, EventArgs e)
        {
            fpSpread2.Open("\\\\mvfile\\MV REPORTS\\Master Data & Forms\\PlanningExecutionForms\\VCAPP\\ShipmentTemplate.XML");
            //Cursor.Current = Cursors.WaitCursor;
            if (cmbVendorLoadShip.Text.ToString().Trim() != "SELECT")
            {
                string vendorname = cmbVendorLoadShip.Text.ToString();
                string folderName = getFolderName(vendorname);
                fname = "Shipment Tracker " + vendorname;
                //    fname = "Test Shipment Tracker Intel";
                //    //****************get FileID from share doc **************************************************/

                string pageToken = null;
                do
                {

                    string parent_id = getGIDByName(folderName);
                    var result = getFilesByParentId(parent_id, pageToken);
                    foreach (var file in result.Files)
                    {

                        if (file.Name == fname)
                        {
                            ListViewItem item = new ListViewItem(new string[2] { file.Name, file.Id });
                            fileLink = file.Id;
                            break;
                        }
                    }
                    pageToken = result.NextPageToken;
                } while (pageToken != null);

                //    fileLink = "1IEVoJrJDmTEqjRTCejz8_Tw2E-rsiadjjA2DB9qLvaw";

                //****************Download File by fileID from share doc **************************************************/
                if (this.fileLink != "")
                {
                    Google.Apis.Drive.v3.Data.File file1 = service.Files.Get(this.fileLink).Execute();
                    saveTo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\ShipTracker" + vendorname + ".xlsx";
                    saveTo = saveTo.Replace("file:\\", "");

                    //  saveTo = "\\\\Mvfile\\mv reports\\Master Data & Forms\\PlanningExecutionForms\\ShipmentTracker\\ShipmentTracker" + vendorname + ".xlsx";
                    if (downloadfile(service, fileLink, saveTo))
                    {

                        openExcelInSpreadSheet(saveTo, this.fpspreadShipTracker);
                        //  System.IO.FileStream fs = System.IO.File(saveTo);

                        //****************Open File in EXCEL component fpSpread **************************************************/  

                        //   saveTo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Shipment Tracker CELESTICA.xlsx";
                        //    saveTo = saveTo.Replace("file:\\", "");
                        /*            Microsoft.Office.Interop.Excel.Worksheet sheet = null;
                                    Microsoft.Office.Interop.Excel.Sheets excelSheets = null;
                                    //Microsoft.Office.Interop.Excel.Application excelApp;
                                    Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
                                    Microsoft.Office.Interop.Excel.Workbooks excelWorkbooks = null;
                                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                                    string workbookPath = saveTo.Replace("file:\\", "");
                                    excelWorkbooks = excelApp.Workbooks;
                                    excelWorkbook = excelWorkbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                                    excelWorkbook.Close(true);
                                    try
                                    {
                                        this.fpspreadShipTracker.OpenExcel(saveTo);
                                    }

                                    catch
                                    {
                                        MessageBox.Show("Please check if file is in the folder  " + saveTo);
                                        goto Nofile;
                                    }*/
                        //    fpSpread3.OpenExcel(saveTo);
                        // }
                        int rowCount = 0;
                        //****************Creating Doc ID **************************************************/
                        DataSet dsVend = getdataSet("    SELECT * from m_app_getVendorPart where  ep_vend_code='" + this.cmbVendorLoadShip.Text.ToString() + "'");
                        for (int k = 1; k < fpspreadShipTracker.Sheets[0].Rows.Count; k++)
                        {
                            if (fpspreadShipTracker.Sheets[0].Cells[k, 2].Value != null)
                            {
                                if (fpspreadShipTracker.Sheets[0].Cells[k, 0].Value == null)
                                {
                                    fpspreadShipTracker.Sheets[0].Cells[k, 0].Value = System.DateTime.Now.Year.ToString().Substring(2, 2) + System.DateTime.Now.DayOfYear.ToString().PadLeft(3, '0') + cmbVendorLoadShip.SelectedValue.ToString() + (k + 1).ToString().PadLeft(3, '0');
                                }
                            }
                            else
                            {
                                rowCount = k - 1;
                                break;
                            }
                        }
                        //****************Save EXCEL COMPONENT data to File  **************************************************/

                        filename = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\testdoc1.xls";
                        filename = filename.Replace("file:\\", "");

                        DataTable dt = fpspreadShipTracker.DataSource as DataTable;
                        fpspreadShipTracker.SaveExcel(filename);

                        //****************MOVE Data to Archive Tab *************************************************************/
                        if (rowCount > 50)
                            moveToArchive(filename, rowCount);
                        //            
                        //****************Upload Excel file to google share doc**************************************************/
                        updateDoc(filename);

                        //****************Load file from google share doc to Excel Component*************************************/
                        addColumn(filename);
                        loadInterface(filename, dsVend);

                    }
                }
            }
        Nofile:
            Cursor.Current = Cursors.Default;
        }


        //****************Load file from google share doc to Excel Component*************************************//
        public void loadInterface(string filename, DataSet dsVend)
        {
            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.DataSet DtSet;
            System.Data.OleDb.OleDbDataAdapter MyCommand;
            MyConnection = new System.Data.OleDb.OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1';");
            MyConnection.Open();
            DataTable dtExcelSchema;
            dtExcelSchema = MyConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
         string sheetName ;//= dtExcelSchema.Rows[1]["TABLE_NAME"].ToString();
            DataTable schemaTable;
            schemaTable = MyConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Columns, null);
            string colname = "";
            for (int l = 0; l < schemaTable.Rows.Count; l++)
            {
                if (!schemaTable.Rows[l][3].ToString().Contains("F") && !colname.Contains(schemaTable.Rows[l][3].ToString()))

                    colname = colname + ",[" + schemaTable.Rows[l][3].ToString() + "]";
            }
            colname = colname.Substring(1, colname.Length - 1);
            MyCommand = new System.Data.OleDb.OleDbDataAdapter("select [DOC ID],[PO NUMBER],[GPN],[SHIP TO],[SHIP DATE],[SHIP QTY],[PACKING SLIP],[SHIP METHOD],[CARRIER],[TRACKING NUM],[ASN (optional)],[COMMENTS],[TRACK] from [Tracker$]", MyConnection);
            MyCommand.TableMappings.Add("Table", "TestTable");
            DtSet = new System.Data.DataSet();
            MyCommand.Fill(DtSet);
            MyConnection.Close();
            DtSet.Tables[0].Columns.Add("SHIP ID");
            for (int k = 0; k < DtSet.Tables[0].Rows.Count; k++)
            {
                if (DtSet.Tables[0].Rows[k]["GPN"].ToString() != "")
                {
                    string shipDate = DtSet.Tables[0].Rows[k]["SHIP DATE"].ToString();
                    string part_num = DtSet.Tables[0].Rows[k]["GPN"].ToString(); //dt.Rows[0]["GPN"].ToString();
                    string site = "";
                    site = DtSet.Tables[0].Rows[k]["SHIP TO"].ToString(); //dt.Rows[0]["Ship To"].ToString();
                    string start_date = DtSet.Tables[0].Rows[k]["SHIP DATE"].ToString(); //dt.Rows[0]["Ship Date"].ToString();
                    string part_id = "";
                    try
                    {
                        part_id = dsVend.Tables[0].Select("part_num='" + part_num + "' or AltPartNum='" + part_num + "'")[0]["vend_part_id"].ToString();
                    }
                    catch (Exception e)
                    {

                    }
                    string vend_id = dsVend.Tables[0].Rows[0]["vend_id"].ToString();
                    DateTime dd = Convert.ToDateTime(start_date);
                    string weeknum = "";
                    // DateTime ddMon = DateTimeExtensions.StartOfWeek(dd, System.DayOfWeek.Monday);
                    string sqlQry = " select top 1 cast(right(myear,2)as varchar)+(REPLICATE('0', 2-LEN(mweek))+ cast(mweek as varchar)) as yrwk,StartDate from m_weekNumber where StartDate= (select top 1 dbo.udf_SetDateToMonOfThatWk('" + start_date + "'))";
                    DataSet dsWeek = getdataSet(sqlQry);

                    try
                    {
                        weeknum = dsWeek.Tables[0].Rows[0][0].ToString();
                    }
                    catch
                    {

                    }


                    string ship_id = "";
                    if (vend_id != "" && part_id != "" && site != "" && weeknum != "")
                        ship_id = vend_id + site + part_id + weeknum;

                    DtSet.Tables[0].Rows[k]["SHIP ID"] = ship_id;
                    DtSet.Tables[0].AcceptChanges();
                    int m = k + 1;
                    this.fpSpread2.Sheets[0].Cells[m, 0].Value = DtSet.Tables[0].Rows[k]["DOC ID"].ToString();
                    fpSpread2.Sheets[0].Cells[m, 1].Value = DtSet.Tables[0].Rows[k]["SHIP ID"].ToString();
                    fpSpread2.Sheets[0].Cells[m, 3].Value = DtSet.Tables[0].Rows[k]["PO NUMBER"].ToString();  //
                    fpSpread2.Sheets[0].Cells[m, 4].Value = DtSet.Tables[0].Rows[k]["GPN"].ToString(); //[GPN]
                    fpSpread2.Sheets[0].Cells[m, 5].Value = DtSet.Tables[0].Rows[k]["SHIP TO"].ToString(); //[SHIP TO]
                    fpSpread2.Sheets[0].Cells[m, 6].Value = DtSet.Tables[0].Rows[k]["SHIP DATE"].ToString();//[SHIP DATE]
                    fpSpread2.Sheets[0].Cells[m, 7].Value = DtSet.Tables[0].Rows[k]["SHIP QTY"].ToString();//[SHIP QTY]
                    fpSpread2.Sheets[0].Cells[m, 8].Value = DtSet.Tables[0].Rows[k]["PACKING SLIP"].ToString();//[PACKING SLIP]
                    fpSpread2.Sheets[0].Cells[m, 9].Value = DtSet.Tables[0].Rows[k]["SHIP METHOD"].ToString();// [SHIP METHOD]
                    fpSpread2.Sheets[0].Cells[m, 10].Value = DtSet.Tables[0].Rows[k]["CARRIER"].ToString();//[CARRIER]
                    string tracking_num = "";
                    double track;
                    try
                    {
                        tracking_num = DtSet.Tables[0].Rows[k]["TRACK"].ToString();
                        fpSpread2.Sheets[0].Cells[m, 11].Value = tracking_num;
                    }
                    catch (Exception eee)
                    {
                        // fpSpread2.Sheets[0].Cells[m, 12].Value = DtSet.Tables[0].Rows[k][12].ToString();
                    }
                    fpSpread2.Sheets[0].Cells[m, 12].Value = DtSet.Tables[0].Rows[k]["ASN (optional)"].ToString();//[ASN (optional)]
                    fpSpread2.Sheets[0].Cells[m, 13].Value = DtSet.Tables[0].Rows[k]["COMMENTS"].ToString();//[COMMENTS]
                    // fpSpread2.Sheets[0].Cells[m, 14].Value = DtSet.Tables[0].Rows[k][2].ToString();//[COMMENTS]
                }
            }
        }
        public void moveToArchive(string filename, int rowCount)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            // Microsoft.Office.Interop.Excel.Workbook xlWb=null;

            Microsoft.Office.Interop.Excel.Workbook xlWb = xlApp.Workbooks.Open(@filename.Trim());//@"E:\\Ruchi\\NewPlanningExceution\\MaintainGoogleDoc\\MaintainGoogleDoc\\bin\\Debug\\testdoc.xls");              
            Microsoft.Office.Interop.Excel.Worksheet xlWs = (Microsoft.Office.Interop.Excel.Worksheet)xlWb.Sheets[1]; // Sheet1            
            Microsoft.Office.Interop.Excel.Worksheet xlWs1 = (Microsoft.Office.Interop.Excel.Worksheet)xlWb.Sheets[2]; // Sheet1        
            try
            {
                if (rowCount > 50)
                {
                    int cnt = rowCount - 49;
                    // cut column B and insert into A, shifting columns right                   
                    Microsoft.Office.Interop.Excel.Range copyRange = xlWs.Range["A2:L" + cnt];
                    Microsoft.Office.Interop.Excel.Range insertRange = xlWs1.Range["A2:L2"];
                    insertRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, copyRange.Cut());
                    xlWs.Range["A2:L" + cnt].Delete();
                }
                Microsoft.Office.Interop.Excel.Range headerRange = xlWs.Range["A1:L100"];
                //    headerRange.Locked = true;
                //   xlWs.Protect();
                xlApp.DisplayAlerts = false;
                xlApp.SaveWorkspace();
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message.ToString());
            }
            finally
            {
                //Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
                xlWb.Close(true, false, Type.Missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            }
        }

        public void updateDoc(string filename)
        {
            filename = filename.Replace("file:\\", " ");
            try
            {
                Google.Apis.Drive.v3.Data.File body = new Google.Apis.Drive.v3.Data.File();
                body.Name = fname;
                body.Description = fname;
                body.MimeType = "application/xlsx";
                System.IO.Stream fileStream = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                byte[] byteArray = new byte[fileStream.Length];
                fileStream.Read(byteArray, 0, (int)fileStream.Length);
                System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
                FilesResource.UpdateMediaUpload request = service.Files.Update(body, this.fileLink, stream, "application/x-vnd.oasis.opendocument.spreadsheet");

                // request.Convert = true;
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

        private void btnSaveShipment_Click(object sender, EventArgs e)
        {
            string strSave = "";

            for (int k = 1; k < this.fpSpread2.Sheets[0].Rows.Count; k++)
            {
                if (fpSpread2.Sheets[0].Cells[k, 2].Value.ToString() == "YES")
                {
                    string docId = "";
                    string shipId = "";
                    string po_num = "";
                    string gpn = "";
                    string site = "";
                    string ship_date = "";
                    int ship_qty = 0;
                    string packing_slip = "";
                    string ship_method = "";
                    string carrier = "";
                    string tracking_num = "";
                    string asn = "";
                    string comments = "";

                    if (fpSpread2.Sheets[0].Cells[k, 0].Value != null)
                        docId = fpSpread2.Sheets[0].Cells[k, 0].Value.ToString();

                    if (fpSpread2.Sheets[0].Cells[k, 1].Value != null)
                        shipId = fpSpread2.Sheets[0].Cells[k, 1].Value.ToString();

                    if (fpSpread2.Sheets[0].Cells[k, 3].Value != null)
                        po_num = fpSpread2.Sheets[0].Cells[k, 3].Value.ToString();

                    if (fpSpread2.Sheets[0].Cells[k, 4].Value != null)
                        gpn = fpSpread2.Sheets[0].Cells[k, 4].Value.ToString();

                    if (fpSpread2.Sheets[0].Cells[k, 5].Value != null)
                        if (fpSpread2.Sheets[0].Cells[k, 5].Value.ToString().Length == 2)
                            site = fpSpread2.Sheets[0].Cells[k, 5].Value.ToString();

                    if (fpSpread2.Sheets[0].Cells[k, 6].Value != null)
                        ship_date = fpSpread2.Sheets[0].Cells[k, 6].Value.ToString();

                    if (fpSpread2.Sheets[0].Cells[k, 7].Value != null)
                        ship_qty = Convert.ToInt32(fpSpread2.Sheets[0].Cells[k, 7].Value.ToString().Replace(",", ""));

                    if (fpSpread2.Sheets[0].Cells[k, 8].Value != null)
                        packing_slip = fpSpread2.Sheets[0].Cells[k, 8].Value.ToString();

                    if (fpSpread2.Sheets[0].Cells[k, 9].Value != null)
                        ship_method = fpSpread2.Sheets[0].Cells[k, 9].Value.ToString();

                    if (fpSpread2.Sheets[0].Cells[k, 10].Value != null)
                        carrier = fpSpread2.Sheets[0].Cells[k, 10].Value.ToString();

                    if (fpSpread2.Sheets[0].Cells[k, 11].Value != null)
                        //try
                        //{
                        //    tracking_num = Convert.ToString(Double.TryParse(fpSpread2.Sheets[0].Cells[k, 11].Value.ToString(), System.Globalization.NumberStyles.Any,));
                        //}
                        //catch
                        //{
                        tracking_num = fpSpread2.Sheets[0].Cells[k, 11].Value.ToString();
                    // }

                    if (fpSpread2.Sheets[0].Cells[k, 12].Value != null)
                        asn = fpSpread2.Sheets[0].Cells[k, 12].Value.ToString();

                    if (fpSpread2.Sheets[0].Cells[k, 13].Value != null)
                        comments = fpSpread2.Sheets[0].Cells[k, 13].Value.ToString();
                    strSave = strSave + "   exec sp_APP_VCA_SaveShipTracker '" + docId + "','" + shipId + "','" + site + "','" + gpn + "','" + ship_date + "','" + ship_qty + "','" + ship_method + "','" + packing_slip + "','" + carrier + "','" + tracking_num + "','" + asn + "','" + comments.Replace("'", "''") + "'";

                }
            }

            if (strSave != "")
            {
                SqlConnection cnsave = new SqlConnection(constr);
                cnsave.Open();
                SqlTransaction tran = cnsave.BeginTransaction();
                try
                {
                    SqlCommand cmd = new SqlCommand(strSave, cnsave);
                    cmd.Transaction = tran;
                    cmd.ExecuteNonQuery();
                    tran.Commit();
                    MessageBox.Show("ShipTracker has been inserted successfully", "SUCESS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    MessageBox.Show("ShipTracker has been failed to save the data " + ex.Message.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    cnsave.Dispose();
                    cnsave.Close();
                }
            }
        }

        public void updateUploadFlag()
        {
            string strSave = "Update [MIMDIST].[dbo].[VCA_ReqCom_Table] set REQ_ReqCreatedFlag='Y',ReqIDDateTime=getdate()  where req_id=" + txtReqID.Value;

            SqlConnection cnsave = new SqlConnection(constr);
            cnsave.Open();

            try
            {
                SqlCommand cmd = new SqlCommand(strSave, cnsave);
                cmd.ExecuteNonQuery();

            }
            catch (Exception e)
            {
            }


        }

        private void cmbVendorRequest_SelectedIndexChanged(object sender, EventArgs e)
        {
            getRequestId();
        }

        private void cmbWeekRequest_SelectedIndexChanged(object sender, EventArgs e)
        {
            getRequestId();
        }


        private void cmbYearRequest_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void cmbVendPart_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void fpspreadShipTracker_CellClick(object sender, FarPoint.Win.Spread.CellClickEventArgs e)
        {

        }
        Microsoft.Office.Interop.Excel.Worksheet sheet;
        Microsoft.Office.Interop.Excel.Sheets excelSheets;
        Microsoft.Office.Interop.Excel.Application excelApp1;
        Microsoft.Office.Interop.Excel.Workbook excelWorkbook;
        private void addColumn(string filename)
        {
            filename = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\testdoc1.xls";
            string workbookPath = filename.Replace("file:\\", "");
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                sheetnum = 0;
                excelSheets = excelWorkbook.Worksheets;
                Microsoft.Office.Interop.Excel.Worksheet sheet = excelSheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                sheet.Cells[1, 13] = "TRACK";
                for (int rownum = 2; rownum < 100; rownum++)
                {
                    try
                    {
                        sheet.Cells[rownum, 13] = "=TRIM(J" + rownum + ")";
                        //   sheet.Cells[rownum, 13] = "=J" + rownum + "";
                    }
                    catch (Exception ex)
                    {
                        string error1 = ex.Message.ToString();
                    }
                }
                excelWorkbook.CheckCompatibility = false;
                excelWorkbook.DoNotPromptForConvert = true;
                excelWorkbook.Save();
            }
            catch (Exception ee)
            {
                string error = ee.Message.ToString();
            }
            finally
            {
                //Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
                try
                {
                    excelWorkbook.Close(true, false, Type.Missing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                catch { }
            }
        }

        private void BtnExportShipment_Click(object sender, EventArgs e)
        {
            DataTable dt = dgvShipTracker.DataSource as DataTable;
            string filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\shiptracker.csv";
            ExportToCSV(dt, filename);

            MessageBox.Show("File has been exported to your desktop " + filename);
            //  OpenCSVWithExcel(path);
            /* try
             {
                 String EXL = "C:\\Program Files\\Microsoft Office\\Office14\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
                 System.Diagnostics.Process proc = new System.Diagnostics.Process();
                 proc.StartInfo.FileName = EXL;
                 proc.StartInfo.Arguments = filename;
                 proc.StartInfo.UseShellExecute = false;
                 proc.StartInfo.RedirectStandardOutput = true;               
                 proc.Start();
             }
             catch
             {
                 String EXL = "C:\\Program Files (x86)\\Microsoft Office\\Office14\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
                 System.Diagnostics.Process proc = new System.Diagnostics.Process();
                 proc.StartInfo.FileName = EXL;
                 proc.StartInfo.Arguments = filename;
                 proc.StartInfo.UseShellExecute = false;
                 proc.StartInfo.RedirectStandardOutput = true;
                 proc.Start();
             }
           
           */


            /*     String EXL = "HKLM\\Software\\Microsoft\\Office\\12.0\\Common\\InstallRoot" + "\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
                  System.Diagnostics.Process proc = new System.Diagnostics.Process();
                  proc.StartInfo.FileName = EXL;
                  proc.StartInfo.Arguments = path;
                  proc.StartInfo.UseShellExecute = false;
                  proc.StartInfo.RedirectStandardOutput = true;
                  proc.Start();*/
        }

        static void OpenCSVWithExcel(string path)
        {
            var ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Workbooks.OpenText(path, Comma: true);
            ExcelApp.Visible = true;
        }

        private void ExportToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers  
            //sw.Write("ID");
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        private void btnForecast_Click(object sender, EventArgs e)
        {
            SqlConnection cn = new SqlConnection(constrMVint);
            cn.Open();
            string sqlQueryE2Open = "  Select * from [dbo].[v_createBTS_Forecast] where [VendorID]= " + cmbVendorE2Open.SelectedValue;
            string sqlForecastId = "  select isnull(tranForecastID,0) from  [PRODDIST].[dbo].[BTS_GoogleIntForecast]  where AutoId =(SELECT max([AutoId])       FROM [PRODDIST].[dbo].[BTS_GoogleIntForecast] where VendorID= '" + cmbVendorE2Open.SelectedValue+"')";
            SqlCommand cmdE2Open = new SqlCommand(sqlQueryE2Open + sqlForecastId, cn);
            SqlDataAdapter daE2Open = new SqlDataAdapter(cmdE2Open);
            DataSet dsE2Open = new DataSet();
            DataTable dtE2Open =new DataTable();
            daE2Open.Fill(dsE2Open);
            if (dsE2Open.Tables.Count > 1)
            {
                dtE2Open = dsE2Open.Tables[0] as DataTable;
                if (dsE2Open.Tables[1].Rows.Count > 0)
                {
                    lblForecastId.Text = dsE2Open.Tables[1].Rows[0][0].ToString();
                    if (lblForecastId.Text.Length > 15)
                        lblForecastDate.Text = lblForecastId.Text.ToString().Substring(2, 4) + "/" + lblForecastId.Text.ToString().Substring(6, 2) + "/" + lblForecastId.Text.ToString().Substring(8, 2) + " " + lblForecastId.Text.ToString().Substring(10, 2) + ":" + lblForecastId.Text.ToString().Substring(12, 2) + ":" + lblForecastId.Text.ToString().Substring(14, 2);
                }
            }
            cmdE2Open.Dispose();
            cn.Close();
           
            dgE2Open.DataSource = dtE2Open;


        }

        private void btnUploadE2Open_Click(object sender, EventArgs e)
        {

            SqlConnection cn = new SqlConnection(constrMVint);
            cn.Open();

            string spE2Open = "  Exec [dbo].[sp_CreateBTS_Forecast]  '" + cmbVendorE2Open.SelectedValue +"'";

            SqlCommand cmdE2Open = new SqlCommand(spE2Open, cn);
            SqlDataAdapter daE2Open = new SqlDataAdapter(cmdE2Open);
        
            DataTable dtE2Open = new DataTable();
            try
            {
                daE2Open.Fill(dtE2Open);
                if (dtE2Open != null)
                    if (dtE2Open.Rows.Count>0)
                    {
                            lblForecastId.Text = dtE2Open.Rows[0][0].ToString();
                         //   lblForecastDate.Text = lblForecastId.Text.ToString().Substring(0, 4) + "/" + lblForecastId.Text.ToString().Substring(5, 2) + "/" + lblForecastId.Text.ToString().Substring(7, 2) + " " + lblForecastId.Text.ToString().Substring(9, 2) + ":" + lblForecastId.Text.ToString().Substring(11, 2);
                    }
                MessageBox.Show("Uploaded Successfully in E2Open");
            }
            catch (Exception ex)
            { 
            
            }
            cmdE2Open.Dispose();
            cn.Close();

           

        }
    }     
       
    }

public static class DateTimeExtensions
{
    public static DateTime StartOfWeek(this DateTime dt, DayOfWeek startOfWeek)
    {
        int diff = dt.DayOfWeek - startOfWeek;
        if (diff < 0)
        {
            diff += 7;
        }

        return dt.AddDays(-1 * diff).Date;
    }
}