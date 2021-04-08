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
using System.Diagnostics;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading.Tasks;
using Google.Apis.Download;

namespace Version3
{
    public partial class RackDeliveryPlan : Form
    {
        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        public RackDeliveryPlan()
        {
            InitializeComponent();
            getauthorization();
        }

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

            string pageToken = null;
            string fileId = "";


            do
            {
                var request = service.Files.List();
                // request.Q = "mimeType = 'application/vnd.google-apps.folder' ";
                string parent_id = "0B1auUUiGcbbBOEIyNjg0alkwc3c";
                request.Q = "'" + parent_id + "' in parents";
                request.Spaces = "Drive";
                request.Fields = "nextPageToken, files(id, name)";
                request.PageToken = pageToken;

                var result = request.Execute();
                foreach (var file in result.Files)
                {
                    if (file.Name == "Rack Delivery Plan [go/rackdeliveryplan]")
                    {
                        fileId = file.Id;
                        break;
                    }
                }
                pageToken = result.NextPageToken;
            } while (pageToken != null);

            string fileName = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\currentRackDeliveryPlan.xlsx";
            fileName = fileName.Replace("file:\\", "");

            Google.Apis.Drive.v3.Data.File file1 = service.Files.Get(fileId).Execute();
            if (this.downloadfile(service, fileId, fileName))
            {
                getallSheet(fileName);
              //  MessageBox.Show("File has been downloaded");
            }
        }


        public bool downloadfile(DriveService service, string fileId, string fileName)
        {
            var request = service.Files.Export(fileId, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            //   var fileId1 = "0BwwA4oUTeiV1UVNwOHItT0xfa2M";
            //  var request = service.Files.Get(fileId1);

             Stream  streamvar=   request.ExecuteAsStream();
             using (FileStream fs = System.IO.File.Create(fileName))
            {
                
                byte[] buffer = new byte[8 * 1024];
                int len;
                while ((len = streamvar.Read(buffer, 0, buffer.Length)) > 0)
                {
                    fs.Write(buffer, 0, len);
                }    
                 return   true;
            } 
           
            return false;
        }
      
    
        public void getallSheet(string FileName)
        {
            saveTo = FileName;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                string workbookPath = FileName.Replace("file:\\", "");
                excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                sheetnum = 0;
                excelSheets = excelWorkbook.Worksheets;

                foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in excelWorkbook.Worksheets)
                {
                    if (worksheet.Name.Contains("Rack Delivery Plan"))
                    {
                        ListItem1 lst = new ListItem1(worksheet.Name, worksheet.Name);
                         lstWorksheet.Items.Add(lst);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }
            finally
            {
                //Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
                excelWorkbook.Close(true, Type.Missing, Type.Missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }

        }

        public void loadGrid()
        {
            //  dgDemandTracker.DataSource=getDataSet().Tables[0];
        }

        public DataSet getDataSet()
        {
            //string strqry = " SELECT  distinct  site_code as SITE, convert(varchar,getdate(),101) as [FILE DATE], convert(varchar,priority_date,101) as [POP DATE], week_num as [POP WEEK], year as [POP YEAR], build_id as [BUILD ID], datacenter as [CLUSTER] , googlePartNumber as [TLA PART NUM],  max( case when comp_inv_m_fg_form_factor='Rack-ASM' then  wp_part_no end ) as [RACK PART NUM],qty as [TLA QTY],sum(case when comp_inv_master_goo_part_num='07000733' then cast(wp_comp_qty * qty as int) else 0 end )as [12-PK QTY] ,sum(case when comp_inv_master_goo_part_num='07000400' then cast( wp_comp_qty * qty as int) else 0  end ) as  [24-PK QTY],sum( case when comp_inv_m_fg_form_factor='Rack-ASM' then cast(qty * wp_comp_qty as int) end ) as [RACK QTY] FROM         dbo.ted_planning_deliverable_view_rev2 AS ted_planning_deliverable_view_rev2 WHERE     (comp_inv_m_fg_form_factor = 'BATTERY-BOX' OR comp_inv_m_fg_form_factor = 'RACK-ASM') AND (comp_inv_m_fg_mach_type <> 'VANILLA') AND (ted_planning_deliverable_view_rev2.priority_date>= '2011-12-26 00:00:00' AND ted_planning_deliverable_view_rev2.priority_date<'2012-03-13 00:00:00')  group by  status, googlePartNumber, qty, site_code,priority_date, week_num, year, datacenter, build_id order by [pop year],[pop week], site,cluster  select * from mim_menlo_demand_tracker_temp ";
            string sqlqry = " exec m_mim_menlo_demand_tracker_data ";
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlqry, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;
        }

        public DataSet getDataSet_query(string strqry)
        {             
             //string strqry = " SELECT  distinct  site_code as SITE, convert(varchar,getdate(),101) as [FILE DATE], convert(varchar,priority_date,101) as [POP DATE], week_num as [POP WEEK], year as [POP YEAR], build_id as [BUILD ID], datacenter as [CLUSTER] , googlePartNumber as [TLA PART NUM],  max( case when comp_inv_m_fg_form_factor='Rack-ASM' then  wp_part_no end ) as [RACK PART NUM],qty as [TLA QTY],sum(case when comp_inv_master_goo_part_num='07000733' then cast(wp_comp_qty * qty as int) else 0 end )as [12-PK QTY] ,sum(case when comp_inv_master_goo_part_num='07000400' then cast( wp_comp_qty * qty as int) else 0  end ) as  [24-PK QTY],sum( case when comp_inv_m_fg_form_factor='Rack-ASM' then cast(qty * wp_comp_qty as int) end ) as [RACK QTY] FROM         dbo.ted_planning_deliverable_view_rev2 AS ted_planning_deliverable_view_rev2 WHERE     (comp_inv_m_fg_form_factor = 'BATTERY-BOX' OR comp_inv_m_fg_form_factor = 'RACK-ASM') AND (comp_inv_m_fg_mach_type <> 'VANILLA') AND (ted_planning_deliverable_view_rev2.priority_date>= '2011-12-26 00:00:00' AND ted_planning_deliverable_view_rev2.priority_date<'2012-03-13 00:00:00')  group by  status, googlePartNumber, qty, site_code,priority_date, week_num, year, datacenter, build_id order by [pop year],[pop week], site,cluster  select * from mim_menlo_demand_tracker_temp ";
            // string sqlqry = " exec m_mim_menlo_demand_tracker_data ";
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(strqry, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;

        }

        System.Collections.Hashtable editUriTable;
      
        string fileName = "Rack Delivery Plan";
        //  string fileName = "BP Tracker";
        string worksheetname;     
        System.Collections.Hashtable editUriTableCell;   
        String ALLColumns = "";
  
        private void btnLoadGrid_Click(object sender, EventArgs e)
        {
            loadGrid();
        }


        private void DataGridView1_CellFormatting(System.Object sender, System.Windows.Forms.DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if ((e.ColumnIndex == 8) && (e.Value.ToString() == ""))
                {
                    e.CellStyle.BackColor = Color.Red;
                }
            }
            catch 
            { 

            }
        }

        DataTable table;
        public string sharedoccolumns = "";
      
        private void dgDownload_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


        }

        string WorkSheet_name = "";

        private void btnGetDataFromDoc_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            loadDropDown();
            if (lstWorksheet.SelectedItems.Count > 0)
            {
                ListItem1 lst = lstWorksheet.SelectedItem as ListItem1;
                WorkSheet_name = lst.Name.ToString();
                showInGrid(WorkSheet_name);
            }  
            Cursor.Current = Cursors.Default;
        }
        
     
        string fileid ="";
        string Parentid="" ;
        string Collection = "MIM Menlo Demand Tracker";
        string docName = "Rack Delivery Plan [go/rackdeliveryplan]";
        string saveTo = "";
        //public void ImportsharedocToExcel()
        //{
        //    UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
        //          new ClientSecrets
        //          {
        //              ClientId = "679763664590-5tugcvfsp90pbc707t4f8dlma3oj9bab.apps.googleusercontent.com",
        //              ClientSecret = "2lneARBZtpxrbX9vVkxcDiXe",
        //          },
        //          new[] { DriveService.Scope.Drive },
        //          "user", CancellationToken.None).Result;
        //    service = new DriveService(new BaseClientService.Initializer()
        //    {
        //        HttpClientInitializer = credential,
        //        ApplicationName = "MIM Project ",
        //    });
           
        //    FilesResource.ListRequest lstreq = service.Files.List();
        //    do
        //    {
        //        try
        //        {
        //            lstreq.Q = " title ='" + Collection + "'";
        //            FileList fileList = lstreq.Execute();
        //            foreach (Google.Apis.Drive.v2.Data.File fileitem in fileList.Items)
        //            {
        //                //if (fileitem.Title == Collection)//"MIM NPI:" + projectName)
        //               // {   // titlestr = titlestr + "\n " + parentitem.Title;
        //                    Parentid = fileitem.Id;
        //                    break;
        //              //  }
        //            }
        //            lstreq.PageToken = fileList.NextPageToken;
        //        }
        //        catch (Exception e)
        //        {
        //            Console.WriteLine("An error occurred: " + e.Message);
        //            lstreq.PageToken = null;
        //        }
        //    } while (!String.IsNullOrEmpty(lstreq.PageToken));
        //  fileid = getchildInFolder(service, Parentid);
        //}

        //string saveTo = "";
        //public string getchildInFolder(DriveService service, String folderId)
        //{
        //    ChildrenResource.ListRequest request = service.Children.List(folderId);          
        //    do
        //    {
        //        try
        //        {
        //            ChildList children = request.Execute();
        //            foreach (ChildReference child in children.Items)
        //            {
        //                Google.Apis.Drive.v2.Data.File file1 = service.Files.Get(child.Id).Execute();
        //                if (file1.Title == "Rack Delivery Plan [go/rackdeliveryplan]")
        //                {
        //                    fileid = file1.Id;
        //                    saveTo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Templates\\currentRackDeliveryPlan.xlsx";
        //                    saveTo = saveTo.Replace("file:\\", "");
        //                    System.IO.FileStream fs = new System.IO.FileStream(saveTo, System.IO.FileMode.OpenOrCreate);
        //                    fs.Close();
        //                    if (downloadFile(service, file1, saveTo))
        //                        getallSheet(saveTo);
        //                       // showInGrid();
        //                }
        //            }
        //            request.PageToken = children.NextPageToken;
        //        }
        //        catch (Exception e)
        //        {
        //            Console.WriteLine("An error occurred: " + e.Message);
        //            request.PageToken = null;
        //        }
        //    } while (!String.IsNullOrEmpty(request.PageToken));

        //    return fileid;
        //}


        //public void getallSheet(string FileName)
        //{
        //    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();    
        //    try
        //    {                     
        //        string workbookPath = FileName.Replace("file:\\", "");
        //        excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);         
        //        sheetnum = 0;
        //        excelSheets = excelWorkbook.Worksheets;
           
        //        foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in excelWorkbook.Worksheets)
        //        {
        //           if (worksheet.Name.Contains("Rack Delivery Plan"))
        //            {
        //                ListItem1 lst = new ListItem1(worksheet.Name, worksheet.Name);
        //                lstWorksheet.Items.Add(lst);
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message.ToString());
        //    }
        //    finally
        //    {
        //        //Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
        //        excelWorkbook.Close(true, Type.Missing, Type.Missing);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        //    }
     
        //}

        //public Boolean downloadFile(DriveService _service, Google.Apis.Drive.v3.Data.File _fileResource, string _saveTo)
        //{
        //    if (!String.IsNullOrEmpty(_fileResource.ExportLinks["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]))
        //    {
        //        try
        //        {
        //            var x = _service.HttpClient.GetByteArrayAsync(_fileResource.ExportLinks["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]);
        //            byte[] arrBytes = x.Result;
        //            System.IO.File.WriteAllBytes(_saveTo, arrBytes);
        //            return true;
        //        }
        //        catch (Exception e)
        //        {
        //            Console.WriteLine("An error occurred: " + e.Message);
        //            return false;
        //        }
        //    }
        //    else
        //    {
        //        // The file doesn't have any content stored on Drive.
        //        return false;
        //    }
        //}

        string table_name="" ;
        public void showInGrid(string table_name)
        {
            //  table_name = "Rack Delivery Plan";
           // saveTo="E:\\Ruchi\\PlanningExecutionNew\\DriveApiApp\\Version3\\bin\\Debug\\Rack Delivery Plan Template.xlsx";
            System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + saveTo + "';Extended Properties=Excel 12.0;");
            System.Data.OleDb.OleDbDataAdapter MyCommand = new System.Data.OleDb.OleDbDataAdapter("select  * from ["+table_name+"$]", MyConnection);
            MyCommand.TableMappings.Add("Table", "TestTable");
            System.Data.DataSet DtSet = new System.Data.DataSet();
            MyCommand.Fill(DtSet);
            DataTable dtShareColl = DtSet.Tables[0];
            dtShareColl.Rows[0].Delete();
            MyConnection.Close();
            DataTable dtMapping = getDataSet_query("select [id] ,[field_name] ,[doc_header_alpha] ,[field_type] ,[seq],table_name from m_mim_menlo_dem_track_header_mapping_new where tab_name='" + table_name + "'  order by seq  ").Tables[0];
            DataTable DtGrid = new DataTable();
            ALLColumns = "";

            foreach (DataRow dr in dtMapping.Rows)
            {
                ALLColumns = ALLColumns + dr["field_name"].ToString() + ",";
                DataColumn dc = new DataColumn();
                dc.ColumnName = dr["field_name"].ToString();
                dc.DataType = System.Type.GetType(dr["field_type"].ToString());
                DtGrid.Columns.Add(dc);
                table_name = dr["table_name"].ToString();
            }
            int ind = 0;
            foreach (DataRow drShare in dtShareColl.Rows)
            {
                ind = ind + 1;
                if (ind > 2)
                {
                    DataRow drGrid = DtGrid.NewRow();
                    foreach (DataRow drMapping in dtMapping.Rows)
                    {
                        if (drShare[drMapping["doc_header_alpha"].ToString()].ToString() != "")
                        {
                            string col_name = drMapping["doc_header_alpha"].ToString();
                          
                            string str = drShare[drMapping["doc_header_alpha"].ToString()].ToString();
                            string strFld=drMapping["field_name"].ToString();
                            if (drShare[drMapping["doc_header_alpha"].ToString()].ToString().Trim() != "")
                            {
                                try
                                {
                                    drGrid[drMapping["field_name"].ToString()] = str;// drShare[drMapping["doc_header_alpha"].ToString()].ToString();
                                }
                                catch 
                                { 
                                
                                }
                            }
                        }
                    }
                
                    if (drGrid["build_id"].ToString().Trim() != "")
                    DtGrid.Rows.Add(drGrid);
                }
            }

            dgDownload.DataSource = DtGrid;
           Infragistics.Win.UltraWinGrid.UltraGridBand band = dgDownload.DisplayLayout.Bands[0] as Infragistics.Win.UltraWinGrid.UltraGridBand;
            band.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
            foreach (DataColumn dc in DtGrid.Columns)
            {
                if (dc.ColumnName == "mim_slip_reason_cd")
                {
                    band.Columns["mim_slip_reason_cd"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
                    band.Columns["mim_slip_reason_cd"].ValueList = ddMIMSlipCode;
                }
                if (dc.ColumnName == "mim_status")
                {
                    band.Columns["mim_status"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
                    band.Columns["mim_status"].ValueList = ddMIMStatus;
                }

                if (dc.DataType == System.Type.GetType("System.Int32"))
                {
                    // band.Columns[dc.ColumnName].MaskInput = "n,nnn";
                    // band.Columns[dc.ColumnName].UseEditorMaskSettings = true;
                    band.Columns[dc.ColumnName].MaskInput = "n,nnn";
                    band.Columns[dc.ColumnName].MaskClipMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals;
                    band.Columns[dc.ColumnName].MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeBoth;
                    band.Columns[dc.ColumnName].MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw;
                    band.Columns[dc.ColumnName].PromptChar = ' ';
                }
                else if (dc.DataType == System.Type.GetType("System.DateTime"))
                {
                    band.Columns[dc.ColumnName].MaskInput = "mm-dd-yyyy";
                    //   band.Columns[dc.ColumnName].UseEditorMaskSettings = true;
                    //    band.Columns[dc.ColumnName].MaskInput = "mm-dd-yyyy";
                    band.Columns[dc.ColumnName].MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals;
                  
                }
                lblNumOfRec.Text = dgDownload.Rows.Count.ToString();              
            }
        }


        public void loadDropDown()
        {
            DataTable dtMIMStatus = new DataTable();
            dtMIMStatus.Columns.Add("MIM Status");
            dtMIMStatus.Rows.Add("At-risk");
            dtMIMStatus.Rows.Add("On-track");
            dtMIMStatus.Rows.Add("Build-to-Stock");
            dtMIMStatus.Rows.Add("MIM-Shipped");
            dtMIMStatus.Rows.Add("Pull-from-Stock");
            dtMIMStatus.Rows.Add("Pending Kit");
            ddMIMStatus.DataSource = dtMIMStatus;
            ddMIMSlipCode.DataSource = getDataSet("SELECT [mim_slip_code],[slip_desc]  FROM [m_mim_menlo_dem_slip_cd_master] order by [sort_ord] ").Tables[0];
            ddMIMSlipCode.DisplayMember = "slip_desc";
            ddMIMSlipCode.ValueMember = "mim_slip_code";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Google.GData.Documents.DocumentsService docService;  docService = new Google.GData.Documents.DocumentsService("docService");
            //docService.setUserCredentials("tedatmim@gmail.com","mimi100");
            //string srcFileName = txtFilename.Text;
            //int chrInd = txtFilename.Text.LastIndexOf("\\");
            //string destFileName = txtFilename.Text.Substring(chrInd + 1, (txtFilename.Text.Length) - (chrInd + 1));
            //Google.GData.Documents.DocumentEntry newEntry = docService.UploadDocument(srcFileName, destFileName);
            //MessageBox.Show(newEntry.AlternateUri.ToString() + " Document has been uploaded successfully ");
            //// updateAllCell();
        }

        private void btnUpdateTable_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            ListItem1 lst = lstWorksheet.SelectedItem as ListItem1;
            DataTable dt = dgDownload.DataSource as DataTable;
            string sqlval = "    ";
         
                string sqlIns = "  exec  sp_mim_menlo_demand_tracker ";

                //  demand_source,   build_id, tla_part_num, tla_qty, del_product, del_platform, del_part_num, del_qty, menlo_fill_qty, menlo_status,final_destination,plan_model,order_number,comments) values ";
                //  build_id,rack_part_num,menlo_battery_12pack, menlo_battery_24pack,menlo_battery_juicebox,menlo_battery_PSU_boat, menlo_configured_rack,population_complete,mim_notes,menlo_notes) values ";
                //  string sqlUpd=" update [m_mim_menlo_demand_tracker_temp] set tab_name='" + Machine Rack Plan + "',"

                int v_menlo_filled_qty = 0;
                string sqlTot = "";
                string[] ColArr = ALLColumns.Split(',');
                int numRec = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sqlval = "";
                    bool newData = false;
                    if (dt.Rows[i]["date_updated"].ToString().Trim() != "")
                    {
                        numRec = numRec + 1;
                        String dateUpdate = Convert.ToDateTime(dt.Rows[i]["date_updated"].ToString()).ToShortDateString();
                        String current_date = System.DateTime.Now.ToShortDateString();
                        //  if (dateUpdate == current_date)
                        //  {
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            //  string build_id=dt.Rows[i]["build_id"].ToString();
                            //  string del_part_num=dt.Rows[i]["del_part_num"].ToString();
                            if (dt.Columns[j].ColumnName == "menlo_days_late" || dt.Columns[j].ColumnName == "menlo_fill_qty" || dt.Columns[j].ColumnName == "del_qty" || dt.Columns[j].ColumnName == "tla_qty")
                                sqlval = sqlval + "'" + dt.Rows[i][j].ToString().Replace(",", "") + "', ";
                            else if (dt.Columns[j].ColumnName == "menlo_req_arrival_date" || dt.Columns[j].ColumnName == "menlo_est_arrival_date" || dt.Columns[j].ColumnName == "req_ship_date" || dt.Columns[j].ColumnName == "menlo_ship_complete_date" || dt.Columns[j].ColumnName == "menlo_actual_ship_date" || dt.Columns[j].ColumnName == "menlo_actual_ship_date")
                                if (dt.Rows[i][j].ToString() == "")
                                    sqlval = sqlval + "NULL, ";
                                else
                                    sqlval = sqlval + "'" + dt.Rows[i][j].ToString().Replace(",", "") + "', ";

                            else
                                sqlval = sqlval + "'" + dt.Rows[i][j].ToString().Replace("'", "''") + "', ";
                        }

                        sqlTot = sqlTot + sqlIns + sqlval.Substring(0, sqlval.Length - 2);
                    }
                    //   }
                }
                sqlval = sqlTot;
           // }
            //else
            //{
            //    string sqlIns = "   insert into [m_mim_menlo_demand_tracker_temp](" + this.ALLColumns.Substring(0, ALLColumns.Length - 1) + ")  values('";// +WorkSheet_name + "',";//demand_source,   build_id, tla_part_num, tla_qty, del_product, del_platform, del_part_num, del_qty, menlo_fill_qty, menlo_status,final_destination,plan_model,order_number,comments) values ";
            //    //build_id,rack_part_num,menlo_battery_12pack, menlo_battery_24pack,menlo_battery_juicebox,menlo_battery_PSU_boat, menlo_configured_rack,population_complete,mim_notes,menlo_notes) values ";
            //    int v_menlo_filled_qty = 0;
            //    string sqlTot = "";
            //    string[] ColArr = ALLColumns.Split(',');
            //    for (int i = 0; i < dt.Rows.Count; i++)
            //    {
            //        sqlval = "";
            //        for (int j = 0; j < dt.Columns.Count; j++)
            //        {
            //            if (dt.Columns[j].ColumnName == "menlo_days_late" || dt.Columns[j].ColumnName == "menlo_fill_qty" || dt.Columns[j].ColumnName == "del_qty" || dt.Columns[j].ColumnName == "tla_qty")
            //                sqlval = sqlval + "'" + dt.Rows[i][j].ToString().Replace(",", "") + "', ";
            //            else if (dt.Columns[j].ColumnName == "menlo_req_arrival_date" || dt.Columns[j].ColumnName == "menlo_est_arrival_date" || dt.Columns[j].ColumnName == "req_ship_date" || dt.Columns[j].ColumnName == "menlo_ship_complete_date" || dt.Columns[j].ColumnName == "menlo_actual_ship_date" || dt.Columns[j].ColumnName == "menlo_actual_ship_date")
            //                if (dt.Rows[i][j].ToString() == "")
            //                {
            //                    sqlval = sqlval + "NULL, ";
            //                }
            //                else
            //                {
            //                    sqlval = sqlval + "'" + dt.Rows[i][j].ToString().Replace(",", "").Trim() + "', ";
            //                }
            //            else
            //                sqlval = sqlval + "'" + dt.Rows[i][j].ToString().Replace("'", "''").Trim() + "', ";
            //        }
            //        sqlTot = sqlTot + sqlIns + sqlval.Substring(0, sqlval.Length - 2) + ")";
            //    }
            //    sqlval = " delete from [m_mim_menlo_demand_tracker_temp] where tab_name='" + WorkSheet_name + "'" + sqlTot;//where demand_source='" + lst.Name.ToString() + "'" + sqlval;
            //    // else
            //    //  sqlval = " delete from m_mim_menlo_demand_tracker_temp where not demand_source like 'GIG%'" + sqlval;
            //    // MessageBox.Show(sqlval);
            //       }
            
                SqlTransaction tran;
                SqlConnection sqlcon = new SqlConnection(constr);

                sqlcon.Open();
                tran = sqlcon.BeginTransaction();
                SqlCommand cmd = new SqlCommand(sqlval, sqlcon, tran);
                try
                {
                    int cnt = cmd.ExecuteNonQuery();
                    tran.Commit();
                    MessageBox.Show(numRec + " records updated succesfully");
                }
                catch (Exception exp)
                {
                    tran.Rollback();
                    MessageBox.Show("Error to save.Please check!" + exp.Message.ToString());
                }
                sqlcon.Close();
                Cursor.Current = Cursors.Default;            
        }

  
        string col_name;
   

        private void tabUploadtodoc_Click(object sender, EventArgs e)
        {
            DataTable dt = dgDownload.DataSource as DataTable;
            string sqlIns = "   insert into m_mim_menlo_demand_tracker(build_id, tla_part_num, tla_qty, del_product, del_platform, del_part_num, del_qty, menlo_fill_qty, menlo_status) values ";
            //build_id,rack_part_num,menlo_battery_12pack, menlo_battery_24pack,menlo_battery_juicebox,menlo_battery_PSU_boat, menlo_configured_rack,population_complete,mim_notes,menlo_notes) values ";
            string sqlval = "    ";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string v_cust_ord_num = dt.Rows[i]["BUILD ID"].ToString();
                string v_tla_part_num = dt.Rows[i]["TLA Part #"].ToString();
                string v_tla_qty = dt.Rows[i]["TLA QTY"].ToString();
                string v_del_prod = dt.Rows[i]["Deliverable Product"].ToString();
                string v_del_plat = dt.Rows[i]["Deliverable Platform"].ToString();
                string v_del_part_num = dt.Rows[i]["Deliverable Part #"].ToString();
                string v_del_qty = dt.Rows[i]["Deliverable QTY"].ToString();
                string v_menlo_filled_qty = dt.Rows[i]["Menlo fullfilled QTY"].ToString();
                string v_menlo_status = dt.Rows[i]["Menlo Status"].ToString();
                string v_comments = dt.Rows[i]["Comments"].ToString();
                if (v_cust_ord_num != "")
                {
                    sqlval = sqlval + sqlIns + " ('" + v_cust_ord_num + "','" + v_tla_part_num + "','";
                    sqlval = sqlval + v_tla_qty + "','" + v_del_prod + "','" + v_del_plat + "','" + v_del_part_num + "','" + v_del_qty + "','" + v_menlo_filled_qty + "','" + v_menlo_status + "')";
                }
            }

            sqlval = " delete from m_mim_menlo_demand_tracker " + sqlval;
            // MessageBox.Show(sqlval);
            SqlTransaction tran;
            SqlConnection sqlcon = new SqlConnection(constr);

            sqlcon.Open();
            tran = sqlcon.BeginTransaction();
            SqlCommand cmd = new SqlCommand(sqlval, sqlcon, tran);
            try
            {
                cmd.ExecuteNonQuery();
                tran.Commit();
                MessageBox.Show("Saved succesfully");
            }
            catch (Exception exp)
            {
                tran.Rollback();
                MessageBox.Show("Error to save.Please check!" + exp.Message.ToString());
            }
            sqlcon.Close();
        }

        /*   private void DoListQuery()
           {
               // Set up the query
                    ListQuery query = new ListQuery(this.listLink.HRef.Content);
                  //   query.OrderByColumn = this.listOrderByTextBox.Text;
                 //    query.Reverse = this.listQueryReverseCheckBox.Checked;
                  //   query.SpreadsheetQuery = this.listQueryQueryTextBox.Text;

               // Set the view with the new feed
               this.SetListListView(service.Query(query));
           }
   */
        /*      private void DoListInsert()
              {
                  String[] commaSplit; //this.listAddTextBox.Text.Split(',');
                  ListEntry entry = new ListEntry();

                  for (int i=0; i < commaSplit.Length; i++)
                  {
                      String[] pair = commaSplit[i].Split('=');
                      ListEntry.Custom custom = new ListEntry.Custom();
                      custom.LocalName = pair[0];
                      custom.Value = pair[1];
                      entry.Elements.Add(custom);
                  }

           //       AtomEntry retEntry = service.Insert(new Uri(this.worksheetListView.SelectedItems[0].SubItems[2].Text), entry);
              }

         */
        //private void DoListUpdate()
        //   {
        //    string[] colNameArr = col_name.Split(',');
        //    string listval ="";
        //    DataTable dt = this.dgDemandTracker.DataSource as DataTable;
        //   // try
        //    //  {
        //        //   MessageBox.Show(dt.Rows.Count.ToString());
        //    AtomEntry retEntry=new AtomEntry();
        //          for (int k = 0; k < dt.Rows.Count; k++)
        //          {
        //              for (int j = 0; j < colNameArr.Length; j++)
        //              {
        //                  if (colNameArr[j] != "" || colNameArr[j]!=null)
        //                      if (j < dt.Columns.Count)
        //                      {
        //                          if (j + 1 == 18)
        //                          {
        //                            //  listval = listval + colNameArr[j + 1] + ":" + "=Query(IF (AND(((K" + (k + 5) + "+L" + (k + 5) + "+M" + (k + 5) + ")=(N" + (k + 5) + "+O" + (k + 5) + "+P" + (k + 5) + ")) , Q" + (k + 5) + "=1), \"DONE\",\"OPEN\"));";
        //                          }
        //                          else
        //                          {
        //                              listval = listval + colNameArr[j + 1] + ":" + dt.Rows[k][j].ToString() + ";";
        //                          }
        //                      }
        //              }
        //             // listval = listval.Substring(0, listval.Length - 2)+";STATUS:=Query(IF (AND(((K" + (k + 5) + "+L" + (k + 5) + "+M" + (k + 5) + ")=(N" + (k + 5) + "+O" + (k + 5) + "+P" + (k + 5) + ")) , Q" + (k + 5) + "=1), \"DONE\",\"OPEN\"));";
        //              String[] commaSplit = listval.Split(';');
        //              String key =  (k + 5).ToString() ;
        //              ListEntry entry = (ListEntry)(editUriTable[key]);
        //              string oper = "";
        //              if (entry == null)
        //              {
        //                  oper = "Insert";
        //                   entry = new ListEntry();
        //              }
        //                  else
        //              {
        //                  oper = "Update";

        //              }
        //              for (int i = 0; i < commaSplit.Length; i++)
        //              {
        //                  String[] pair = commaSplit[i].Split(':');

        //                  // Make the new replacement custom
        //                  ListEntry.Custom custom = new ListEntry.Custom();
        //                  custom.LocalName = pair[0];
        //                  if (pair.Length == 2)
        //                  {
        //                      custom.Value = pair[1];
        //                  }
        //                  else
        //                  {
        //                      custom.Value = null;
        //                      if (entry != null)
        //                      {
        //                          if (entry.Elements.Contains(custom))
        //                          {
        //                              entry.Elements.Remove(custom);
        //                          }
        //                          else
        //                          {
        //                              entry.Elements.Add(custom);
        //                          }
        //                      }
        //                      continue;
        //                  }

        //                  // Set the replacement custom
        //                  if (oper == "Insert")
        //                  {
        //                      entry.Elements.Add(custom);
        //                  }
        //                  else if (oper == "Update")
        //                  {
        //                      if (entry != null)
        //                      {
        //                          int index = entry.Elements.IndexOf(custom);
        //                          if (index >= 0)
        //                              entry.Elements[index] = custom;
        //                      }
        //                  }
        //              }

        //              if (entry != null)
        //              {
        //                  if (oper == "Insert")
        //                  {
        //                      //retEntry = service.Insert(new Uri(this.listLink.HRef.Content), entry);

        //                      ListEntry insertedRow = lstFeed.Insert(entry) as ListEntry;

        //                      //MessageBox.Show(retEntry.Updated.Minute.ToString());


        //                  }
        //                  else if (oper == "Update")
        //                  {
        //                       //retEntry = service.Update(entry);
        //                       retEntry= entry.Update();
        //                      // MessageBox.Show(retEntry.Updated.Minute.ToString());

        //                  }
        //              }


        //      /*        ListEntry entry1 = (ListEntry)(editUriTable[key]);
        //              ListEntry.Custom custom1 = new ListEntry.Custom();


        //               custom1.LocalName = "status";

        //               custom1.Value = "=IF (AND(((K" + (key + 4) + "+L" + (key + 4) + "+M" + (key + 4) + ")=(N" + (key + 4) + "+O" + (key + 4) + "+P" + (key + 4) + ")) , Q" + (key + 4) + "=1), \"DONE\",\"OPEN\")";

        //                       // Set the replacement custom
        //                  if (entry1 != null)
        //                  {
        //                      int index = entry1.Elements.IndexOf(custom1);
        //                      if (index >= 0)
        //                          entry1.Elements[index] = custom1;

        //                      if (entry1 != null)
        //                      {

        //                          AtomEntry retEntry1 = service.Update(entry1);
        //                      }
        //                  }
        //             */

        //         }
        //        //  }

        //  //    }
        //    /*  catch( Exception e)
        //      {
        //          MessageBox.Show("exception" + e.Message);

        //      }*/


        //    }
        //private void DoListDelete()
        //{
        //    foreach (string key in editUriTable.Keys)
        //    {
        //        if (key != "1" && key != "2" && key != "3")
        //        {
        //            //  try
        //            // {

        //            ListEntry entry = (ListEntry)(editUriTable[key]);
        //            entry.Delete();
        //            //service.Delete((AtomEntry)(this.editUriTable[key]));
        //            // }
        //            //  catch
        //            //  { 
        //            //   }
        //        }
        //    }
        //}


        //private void btnfile_Click(object sender, EventArgs e)
        //{
        //    if (openFileDialog1.ShowDialog() == DialogResult.OK)
        //    {
        //        txtFilename.Text = openFileDialog1.FileName;
        //    }
        //}

        //private void dgDemandTracker_CellContentClick(object sender, DataGridViewCellEventArgs e)
        //{

        //}

        //private void button1_Click(object sender, EventArgs e)
        //{

        //    DoCellUpdate();
        //    MessageBox.Show("Status has been updated");
        //}

        //private void button7_Click(object sender, EventArgs e)
        //{
        //    Cursor.Current = Cursors.WaitCursor;
        //    connectSharedDoc();
        //  /*  if (this.lslWorksheet1.SelectedItems.Count > 0)
        //    {
        //        ListItem1 lst = lslWorksheet1.SelectedItem as ListItem1;
        //        ListQuery query = new ListQuery(lst.ID.ToString());
        //        DataTable dt = DownloadFileToDataGrid(query, lst.Name.ToString());
        //        dgMIMMenloData.DataSource = dt;
        //        dgMIMMenloData.Columns[0].Width = 80;
        //        dgMIMMenloData.Columns[1].Width = 80;
        //        dgMIMMenloData.Columns[2].Width = 80;
        //        dgMIMMenloData.Columns[3].Width = 85;
        //        dgMIMMenloData.Columns[4].Width = 80;
        //        dgMIMMenloData.Columns[5].Visible = false;
        //        dgMIMMenloData.Columns[6].Width = 100;
        //        dgMIMMenloData.Columns[7].Width = 80;
        //        dgMIMMenloData.Columns[8].Width = 80;
        //        dgMIMMenloData.Columns[9].Width = 80;
        //        dgMIMMenloData.Columns[10].Width = 150;
        //        dgMIMMenloData.Columns[11].Width = 150;
        //        dgMIMMenloData.Columns[12].Width = 150;
        //        dgMIMMenloData.Columns[13].Width = 150;
        //        dgMIMMenloData.Columns[14].Width = 150;
        //        dgMIMMenloData.Columns[15].Width = 150;
        //        dgMIMMenloData.Columns[16].Width = 150;
        //    }*/
        //    Cursor.Current = Cursors.Default;
        //}

        /*   private void button6_Click(object sender, EventArgs e)
              {
                Cursor.Current = Cursors.WaitCursor;
                  string strsql = "Insert into m_mim_menlo_demand_tracker(site_code, file_date, pop_date, week_num, year, build_id, cluster, tla_pn, tla_qty, del_product, del_platform, del_part_num, del_qty,menlo_fill_qty,demand_source) values ";
                  DataTable dtMIMMenlo = dgMIMMenloData.DataSource as DataTable;
                  ListItem1 lst = this.lslWorksheet1.SelectedItem as ListItem1;


                  string sqlval = "    ";
                  string v_site ="";
                  string v_file_date = "";
                  string v_date = "";
                  string v_wk ="";
                  string v_yr = "";
                  string v_demand_type = "";
                  string v_build_id = "";
                  string v_cluster ="";
                  string v_tla_part_num = "";
                  int v_tla_qty = 0;
                  string v_del_prod = "";
                  string v_del_plat = "";
                  string v_del_part_num = "";
                  int v_del_qty = 0;
               //   string v_menlo_filled_qty = "";
                  string v_menlo_status = "";
                  string v_comments = "";

                  int v_menlo_filled_qty = 0;
                  for (int i = 0; i < dtMIMMenlo.Rows.Count; i++)
                  {
               
                      v_site = dtMIMMenlo.Rows[i]["Site"].ToString();
                      v_file_date = dtMIMMenlo.Rows[i]["File Date"].ToString();
                      v_date = dtMIMMenlo.Rows[i]["Date"].ToString();
                      v_wk = dtMIMMenlo.Rows[i]["WK"].ToString();
                      v_yr = dtMIMMenlo.Rows[i]["YR"].ToString();
                      v_demand_type = dtMIMMenlo.Rows[i]["Demand Source"].ToString();
                      v_build_id = dtMIMMenlo.Rows[i]["BUILD ID"].ToString();
                      v_cluster = dtMIMMenlo.Rows[i]["CL"].ToString();
                      v_tla_part_num = dtMIMMenlo.Rows[i]["TLA Part #"].ToString();
                    // v_tla_qty = dtMIMMenlo.Rows[i]["TLA QTY"].ToString();
                     if (dtMIMMenlo.Rows[i]["TLA QTY"].ToString() != "")
                     {
                         v_tla_qty = Convert.ToInt32(dtMIMMenlo.Rows[i]["TLA QTY"].ToString().Replace(",", ""));
                     }
                     v_del_prod = dtMIMMenlo.Rows[i]["Deliverable Product"].ToString();
                     v_del_plat = dtMIMMenlo.Rows[i]["Deliverable Platform"].ToString();
                      v_del_part_num = dtMIMMenlo.Rows[i]["Deliverable Part #"].ToString();
                       //v_del_qty = dtMIMMenlo.Rows[i]["Deliverable QTY"].ToString();


                       if (dtMIMMenlo.Rows[i]["Deliverable QTY"].ToString() != "")
                       {
                           v_del_qty = Convert.ToInt32(dtMIMMenlo.Rows[i]["Deliverable QTY"].ToString().Replace(",", ""));
                       }
                       if (dtMIMMenlo.Rows[i]["Menlo fullfilled QTY"].ToString() != "")
                       {
                           v_menlo_filled_qty = Convert.ToInt32(dtMIMMenlo.Rows[i]["Menlo fullfilled QTY"].ToString().Replace(",", ""));
                       }
                       v_menlo_status = dtMIMMenlo.Rows[i]["Menlo Status"].ToString();
                     v_comments = dtMIMMenlo.Rows[i]["Comments"].ToString();
                     if (v_cluster.Length>0)
                     {
                         v_cluster = v_cluster.Substring(0, 2);
                     }
                      if (v_build_id != "")
                      {
                          sqlval = sqlval + strsql + " ('" + v_site + "','" + v_file_date + "','" + v_date + "','" + v_wk + "','" + v_yr + "','" + v_build_id + "','" + v_cluster + "','" + v_tla_part_num + "','";
                          sqlval = sqlval + v_tla_qty + "','" + v_del_prod + "','" + v_del_plat + "','" + v_del_part_num + "','" + v_del_qty + "','" + v_menlo_filled_qty.ToString() + "','" + v_demand_type + "')";
                      }

                  }

                  // sqlval = " delete from m_mim_menlo_demand_tracker " + sqlval;
                  if (lst.Name.ToString().StartsWith("GIG"))
                      sqlval = " delete from m_mim_menlo_demand_tracker where demand_source='" + lst.Name.ToString() + "'" + sqlval;
                  else
                      sqlval = " delete from m_mim_menlo_demand_tracker where not demand_source like 'GIG%'" + sqlval;

                  // MessageBox.Show(sqlval);
                  SqlTransaction tran;
                  SqlConnection sqlcon = new SqlConnection(constr);

                  sqlcon.Open();
                  tran = sqlcon.BeginTransaction();
                  SqlCommand cmd = new SqlCommand(sqlval, sqlcon, tran);
                  try
                  {
                      cmd.ExecuteNonQuery();
                      tran.Commit();yahoo
                      MessageBox.Show("Saved succesfully");
                  }
                  catch (Exception exp)
                  {
                      tran.Rollback();
                      MessageBox.Show("Error to save.Please check!" + exp.Message.ToString());
                  }
                  sqlcon.Close();
                  Cursor.Current = Cursors.Default;
              }
      */
        private void btnDelete_Click(object sender, EventArgs e)
        {
            //   updateAllCell();
            /*    update_cell(3, 1, " Please wait. Updating..............");
                  try
                  {
                      for (uint r = 3; r <= ((WorksheetEntry)wEntries[0]).RowCount.Count; r++)
                          for (uint c = 1; c <= ((WorksheetEntry)wEntries[0]).ColCount.Count; c++)
                          {
                              if (c != 19)

                                  if ((r == 3 && c == 1))
                                  {}
                                  else
                                  {
                                      update_cell(r, c, "   ");
                                  }
                          }
                  }
                  catch(Exception exp)
                  {
                      MessageBox.Show(exp.Message.ToString());
                  }  */
        }

     //   AtomFeed batchFeed;
        //public void updateAllCell()
        //{
        //    getCellFeed(this.cellsLink.HRef.Content);
        //    update_cell(1, 1, " Please wait. Updating..............");
        //    /*            
        //           for (uint r = 6; r <= ((WorksheetEntry)wEntries[0]).RowCount.Count; r++)
        //           {
        //               for (uint c = 1; c <= ((WorksheetEntry)wEntries[0]).ColCount.Count; c++)
        //               {
        //                   // if (c != 18)
        //                   //  {
        //                   //  update_cell(r, c, "");
        //                   //    update_cell_batch(r, c, " ");
        //                   // }
        //                   string index = "R" + r.ToString()  + "C" + c.ToString();
        //                   DoCellDelete(index);
        //               }
        //           }
        //      */
        //    //cellBatchUpdate(cellFeed,batchFeed);
        //    // query = new CellQuery(this.cellsLink.HRef.Content);
        //    CellQuery query = new CellQuery(this.cellsLink.HRef.Content);
        //    CellFeed cellFeed = service.Query(query);
        //    batchFeed = new AtomFeed(cellFeed);
        //    // cellFeed = service.Query(query);
        //    // batchFeed = new AtomFeed(cellFeed);

        //    DataTable dt = getDataSet().Tables[0];
        //    string data;
        //    for (int j = 0; j < dt.Rows.Count; j++)
        //        for (int i = 0; i < dt.Columns.Count; i++)
        //        {
        //            data = dt.Rows[j][i].ToString();
        //            update_cell_batch(UInt32.Parse((j + 6).ToString()), UInt32.Parse((i + 1).ToString()), data);

        //        }
        //    cellBatchUpdate(cellFeed, batchFeed);
        //    MessageBox.Show("File has been updated sucessfully");
        //    update_cell(1, 1, " UPDATED " + dt.Rows[0][1].ToString());

        //}

        public DataSet getDataSet(string Query)
        {
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(Query, cn);
            cmd.CommandTimeout = 0;
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);            
            cn.Close();
            return ds;
        }

        public void updateGIGSheet()
        {
            String strQuery = " select * from m_mim_menlo_demand_tracker_ship_detail ";
            DataSet dsBPDetail = getDataSet(strQuery);
            int col = 0;

            for (int i = 0; i < dsBPDetail.Tables[0].Rows.Count; i++)
            {
                //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());

                for (int j = 0; j < dsBPDetail.Tables[0].Columns.Count; j++)
                {
                    sheet.Cells[i + 2, j + 1] = dsBPDetail.Tables[0].Rows[i][j].ToString();
                }
            }
            excelWorkbook.Save();
        }

        public void updateBPSheet()
        {
            String strQuery = " exec sp_app_mim_menlo_demand_tracker ";
            DataSet dsBPDetail = getDataSet(strQuery);
            int col = 0;

            // string[] col_share = sharedoccolumns.Split(',');
            for (int j = 0; j < table.Columns.Count; j++)
                sheet.Cells[1, j + 1] = table.Columns[j].ToString();            

            for (int j = 0; j < table.Columns.Count; j++)
                sheet.Cells[2, j + 1] = table.Rows[0][j].ToString();
            
            for (int j = 0; j < table.Columns.Count; j++)
                sheet.Cells[3, j + 1] = table.Rows[1][j].ToString();
            
            for (int i = 0; i < dsBPDetail.Tables[0].Rows.Count; i++)
            {
                //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());
                for (int j = 0; j < dsBPDetail.Tables[0].Columns.Count; j++)
                    sheet.Cells[i + 4, j + 1] = dsBPDetail.Tables[0].Rows[i][j].ToString();                
            }
            excelWorkbook.Save();
        }

        //private void btnUpdate_Click(object sender, EventArgs e)
        //{
        //         create_file();
        //        dr = new Google.Apis.Drive.v2.DriveService(CreateAuthenticator());
        //          if (dr != null)
        //          {
        //              List<Google.Apis.Drive.v2.Data.File> fileList = Google.Apis.Util.Utilities.RetrieveAllFiles(dr);
        //              OpenFileDialog dialog = new OpenFileDialog();
        //              string collectionName = "MIM Menlo Demand Tracker";
        //              string fileName = "MIM-Menlo Demand Tracker";
        //              string Parentid = "";
        //              string archiveParentid = "";
        //              string fileid = "";

        //           foreach (Google.Apis.Drive.v2.Data.File item in fileList)
        //                {
        //                    if (item.Title == "Demand Tracker Archive")//"MIM NPI:" + projectName)
        //                    {
        //                        archiveParentid = item.Id;
        //                        //renameFile(dr, fileid, promptValue + "_Archive");
        //                        break;
        //                    }
        //                }

        //               foreach (Google.Apis.Drive.v2.Data.File item in fileList)
        //              {
        //                  if (item.Title.StartsWith(fileName))//"MIM NPI:" + projectName)
        //                  {
        //                      fileid = item.Id;
        //                      DataSet ds = getDataSet("select  dbo.udf_GetISOWeekNumberFromDate(getdate())");
        //                      string last_week_num = ds.Tables[0].Rows[0][0].ToString();
        //                      File copiedfile = copyFile(dr, fileid, "WeekNum_" + last_week_num+"_"+fileName);                      

        //                    //  renameFile(DriveService service, String fileId, String newTitle)
        //                    //  renameFile(dr,fileid,fileName + "_Archive");
        //                    //  copyFile(dr, fileid, fileName + "_Archive");
        //                    //  insertFileIntoFolder(dr, archiveParentid, copiedfile.Id);
        //                    //  MessageBox.Show(copiedfile.UserPermission.AuthKey.ToString());
        //                      DeleteFile(dr, fileid);
        //                      break;
        //                  }
        //              }

        //               string str="";
        //              foreach (Google.Apis.Drive.v2.Data.File item in fileList)
        //                {
        //                    str = str +  item.Title + " \n";
        //                    if (item.Title.StartsWith("MIM Menlo Demand Tracker"))//"MIM NPI:" + projectName)
        //                    {
        //                        Parentid = item.Id;
        //                        //renameFile(dr, fileid, promptValue + "_Archive");
        //                        break;
        //                    }

        //                }
        //            //  richTextBox1.Text = str;

        //              try
        //              {
        //                File body = new File();
        //                body.Title = System.IO.Path.GetFileName("MIM-Menlo Demand Tracker_" + DateTime.Now.Month.ToString("00") +"_"+ DateTime.Now.Day.ToString("00")+"_" + DateTime.Now.Year.ToString("00") + ".xlsx"); //+ DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + ".xlsx");
        //                  body.Description = "document";
        //                  body.MimeType = "application/xlsx";


        //                  if (!String.IsNullOrEmpty(Parentid))
        //                  {
        //                      body.Parents = new List<ParentReference>() { new ParentReference() { Id = Parentid } };
        //                  }

        //                  filename = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\MIM_Menlo Demand TrackerTemplate.xlsx";// "E:\\Ruchi\\b2bdevMaintainGoogleDoc\\MaintainGoogleDoc\\MaintainGoogleDoc\\bin\\Debug\\MIM_Menlo Demand TrackerTemplate.xlsx";
        //                  filename = filename.Replace("file:\\"," ");
        //                 // MessageBox.Show("filename=" + filename);
        //                  System.IO.Stream fileStream = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read);
        //                  byte[] byteArray = new byte[fileStream.Length];
        //                  fileStream.Read(byteArray, 0, (int)fileStream.Length);
        //                  System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
        //                  FilesResource.InsertMediaUpload request = dr.Files.Insert(body, stream, "application/x-vnd.oasis.opendocument.spreadsheet");

        //                  request.Convert = true;
        //                //  request.Upload();
        //                  MessageBox.Show("File has been uploaded in shared doc");
        //                  fileStream.Close();
        //                  fileStream.Dispose();
        //              }
        //              catch (Exception exp)
        //              {
        //                  MessageBox.Show("Exception to upload file");
        //              }
        //              finally
        //              {

        //              }
        //          }
        //}

        //public static void removeFileFromFolder(DriveService service,  String folderId, String fileId)
        //{
        //    try
        //    {
        //        service.Children.Delete(folderId, fileId).Execute();
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine("An error occurred: " + e.Message);
        //    }
        //}
  /*      public bool insertFileIntoFolder(DriveService service, String folderId, String fileId)
        {


            var getRequest = driveService.Files.Get(fileId);
            var fileMetadata = new Google.Apis.Drive.v3.Data.File();

            fileMetadata.Parents = new List<string> { folderId };
            FilesResource.CreateMediaUpload request;
            using (var stream = new System.IO.FileStream("files/photo.jpg",
                System.IO.FileMode.Open))
            {
                request = file(
                    fileMetadata, stream, "image/jpeg");
                request.Fields = "id";
                request.Upload();
            }
            var file = request.ResponseBody;


            ParentReference newParent = new ParentReference();
            newParent.Id = folderId;
            try
            {
                return service.Parents.Insert(newParent, fileId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }
            return null;
        }
        */
        public static void DeleteFile(DriveService service, String fileId)
        {
            try
            {
                service.Files.Delete(fileId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }
        }
/*
        private static File renameFile(DriveService service, String fileId, String newTitle)
        {
            try
            {
                File file = new File();
                file.Title = newTitle;
                // Rename the file.
                FilesResource.PatchRequest request = service.Files.Patch(file, fileId);
                File updatedFile = request.Execute();

                return updatedFile;
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
                return null;
            }
        }

        private static File copyFile(DriveService service, String originFileId, String copyTitle)
        {
            try
            {
                File copiedFile = new File();
                copiedFile.Title = copyTitle;

                // Rename the file.
                // FilesResource.PatchRequest request = service.Files.Patch(file, fileId);
                // File updatedFile = request.Fetch();

                return service.Files.Copy(copiedFile, originFileId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
                return null;
            }
        }
        */
        private void btnClear_Click(object sender, EventArgs e)
        {

        }

        string filename = "Copy of MIM/Menlo Demand Tracker";

        Microsoft.Office.Interop.Excel.Worksheet sheet;
        Microsoft.Office.Interop.Excel.Sheets excelSheets;
        Microsoft.Office.Interop.Excel.Application excelApp1;
        Microsoft.Office.Interop.Excel.Workbook excelWorkbook;

        int sheetnum = 0;

        string filename1 = "";

        public void create_file()
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            try
            {
                excelApp1 = new Microsoft.Office.Interop.Excel.Application();
                // string workbookPath = "E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\DemandTrackerSheet.xls";// Microsoft.Office.Interop.Excel.XlPlatform.xlWindows  
                string filename = "\\\\mvfile\\MV Reports\\MAster Data & forms\\PlanningExecutionForms\\Templates\\Rack Delivery Plan Template1.xlsx";

                // string filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MIM_Menlo Demand TrackerTemplate.xlsx";
                //filename = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Demand Tracker1_"+DateTime.Now.Month.ToString("00") + "_" + DateTime.Now.Day.ToString("00") + "_" + DateTime.Now.Year.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") +".xlsx";

                filename1 = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Rack Delivery Plan Template.xlsx";
                string workbookPath = filename.Replace("file:\\", "");
                excelWorkbook = excelApp1.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                try
                {
                    excelWorkbook.SaveAs(filename1, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                }
                catch (Exception e)
                {
                    string msg = e.Message.ToString();
                }
              
                //  excelWorkbook = excelApp1.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //   excelWorkbook.SaveAs(filename1, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                //  Microsoft.Office.Interop.Excel.Workbook 
                //  excelWorkbook = this.excelApp1.Workbooks.Add();

                sheetnum = 0;
                excelSheets = excelWorkbook.Worksheets;
                sheetnum = sheetnum + 1;
                sheet = excelSheets[sheetnum] as Microsoft.Office.Interop.Excel.Worksheet;

                String strQuery = "    select   * from dbo.RackTracker_Main order by [POP DATE],[BUILD ID] ";//  select  * from dbo.m_mim_menlo_demand_tracker_main_page_gig order by [ASM DUE DATE],[BUILD ID]  ";//  select  * from m_mim_menlo_demand_tracker_ship_detail  select  * from m_mim_menlo_demand_tracker_gig_status order by [KitOrdered] ";
                DataSet dsDeliveryPlan = getDataSet(strQuery);
                updateMachineRackPlanSheet(dsDeliveryPlan.Tables[0]);
                // updateMachineRackPlanSheetEx(sheet);

             //   sheetnum = sheetnum + 1;
             //   sheet = excelSheets[sheetnum] as Microsoft.Office.Interop.Excel.Worksheet;
             //   updateFabricRackPlanSheet(dsDeliveryPlan.Tables[1]);
                // updateFabricRackPlanSheetEx(sheet);

                sheetnum = sheetnum + 1;
                sheet = excelSheets[sheetnum] as Microsoft.Office.Interop.Excel.Worksheet;
                // updateShipDetailSheet(dsDeliveryPlan.Tables[2]);
                updateShipDetailEx(sheet);

                //sheetnum = sheetnum + 1;
                //sheet = excelSheets[sheetnum] as Microsoft.Office.Interop.Excel.Worksheet;
                ////  updateFabricWIPStatus(dsDeliveryPlan.Tables[3]);
                //updateFabricWIPStatusEx(sheet);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }
            finally
            {
                //Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
                excelWorkbook.Close(true, false, Type.Missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp1);
            }
        }


        private void updateFabricWIPStatusEx(Microsoft.Office.Interop.Excel.Worksheet sheet)
        {
            String stsql1 = " select  * from m_mim_menlo_demand_tracker_gig_status order by [Kit Ordered Date] ";
            string cnt = "data source=mverp;initial catalog=mimdist;user id=sa;password=mimi~100;pooling=true;max pool size=100;min pool size=1;";
            Microsoft.Office.Interop.Excel.QueryTable oQryTable = sheet.QueryTables.Add("OLEDB;Provider=sqloledb;" + cnt, sheet.Range["A1"], stsql1);
            oQryTable.RefreshStyle = Microsoft.Office.Interop.Excel.XlCellInsertionMode.xlInsertEntireRows; // 2; //' xlInsertEntireRows = 2
            oQryTable.Refresh(false);
            excelWorkbook.Save();
        }


        private void updateShipDetailEx(Microsoft.Office.Interop.Excel.Worksheet sheet)
        {
            String stsql1 = " select  * from m_mim_menlo_demand_tracker_ship_detail";
            string cnt = "data source=mverp;initial catalog=mimdist;user id=sa;password=mimi~100;pooling=true;max pool size=100;min pool size=1;";
            Microsoft.Office.Interop.Excel.QueryTable oQryTable = sheet.QueryTables.Add("OLEDB;Provider=sqloledb;" + cnt, sheet.Range["A1"], stsql1);
            oQryTable.RefreshStyle = Microsoft.Office.Interop.Excel.XlCellInsertionMode.xlInsertEntireRows; // 2; //' xlInsertEntireRows = 2
            oQryTable.Refresh(false);
            excelWorkbook.Save();
        }

        public void updateShipDetailSheet(DataTable dt)
        {
            int col = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                }
            }
            excelWorkbook.Save();
        }


        public void updateFabricWIPStatus(DataTable dt)
        {
            int col = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                }
            }
            excelWorkbook.Save();
        }

        public void updateMachineRackPlanSheet(DataTable dt)
        {

            int col = 0;
                int Row = 0;
            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());
                    Row=i;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        col=j;
                        string data = dt.Rows[i][j].ToString();
                        sheet.Cells[i + 4, j + 1] = dt.Rows[i][j].ToString();
                    }
                }
                excelWorkbook.Save();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error  Row number :" + Row.ToString() + " column : " + col.ToString());

                MessageBox.Show("Error" + e.Message.ToString());
            }
        }


        public void updateFabricRackPlanSheet(DataTable dt)
        {
            int col = 0;           
            for (int i = 0; i < dt.Rows.Count; i++)
            {
               for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sheet.Cells[i + 4, j + 1] = dt.Rows[i][j].ToString();
                }
            }
            excelWorkbook.Save();
        }
              

        //public static Google.Apis.Drive.v2.DriveService dr { get; private set; }
        //static String CLIENT_ID = "943908261643.apps.googleusercontent.com";
        //static String CLIENT_SECRET = "63oqKhgS7Gm9I3mx3PBVGcXn";
        //static String REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob http://localhost";
        //static String[] SCOPES = new String[] { "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/userinfo.email", "https://www.googleapis.com/auth/userinfo.profile" };

        private void button1_Click(object sender, EventArgs e)
        {

           Cursor.Current = Cursors.WaitCursor;
         create_file();
         //  FilesResource.ListRequest lstreq = service.Files.List();
          string ArchiveParentid = "";
           string MainParentid = "";

           string pageToken = null;
           string fileId = "";
           do
           {
               var request = service.Files.List();
               // request.Q = "mimeType = 'application/vnd.google-apps.folder' ";
              
               request.Q = "mimeType = 'application/vnd.google-apps.folder'";
               request.Spaces = "Drive";
               request.Fields = "nextPageToken, files(id, name)";
               request.PageToken = pageToken;

               var result = request.Execute();
               foreach (var file in result.Files)
               {                  
                   if (file.Name == "MIM Menlo Demand Tracker")
                   {
                       MainParentid = file.Id;
                       break;
                   }
               }

         
                 foreach (var file in result.Files)
                 {
                     if (file.Name == "Demand Tracker Archive")
                     {
                         ArchiveParentid = file.Id;
                         break;
                     }
                 }
               pageToken = result.NextPageToken;
           } while (pageToken != null);
        
   /*        do
           {         
          var request1 = service.Files.List();
               // request.Q = "mimeType = 'application/vnd.google-apps.folder' ";
              
               request1.Q = "mimeType = 'application/vnd.google-apps.folder'";
               request1.Spaces = "Drive";
               request1.Fields = "nextPageToken, files(id, name)";
               request1.PageToken = pageToken;

               var result1 = request1.Execute();
               foreach (var file in result1.Files)
               {
                   if (file.Name == "Demand Tracker Archive")
                   {
                       ArchiveParentid = file.Id;
                       break;
                   }

                   if (file.Name == "MIM Menlo Demand Tracker")
                   {
                       MainParentid = file.Id;
                       break;
                   }
               }
               pageToken = result1.NextPageToken;
           } while (pageToken != null);
*/
              do
            {
                var request3 = service.Files.List();
                // request.Q = "mimeType = 'application/vnd.google-apps.folder' ";
                string parent_id = "0B1auUUiGcbbBOEIyNjg0alkwc3c";
                request3.Q = "'" + parent_id + "' in parents";
                request3.Spaces = "Drive";
                request3.Fields = "nextPageToken, files(id, name)";
                request3.PageToken = pageToken;

                var result3 = request3.Execute();
                foreach (var file in result3.Files)
                {
                    if (file.Name == "Rack Delivery Plan [go/rackdeliveryplan]")
                    {
                        fileId = file.Id;
                        break;

                    }
                }
                pageToken = result3.NextPageToken;
            } while (pageToken != null);
             // fileId = "1ISM3iI00DMe5WFsHehwdcSf5caWKmazjrubNT6hX7pU";
            moveFile(fileId, ArchiveParentid);
            //  docName = " New Rack Delivery Plan 1115";
                     try
                   {
                       Google.Apis.Drive.v3.Data.File body = new Google.Apis.Drive.v3.Data.File();
                       body.Name = docName;
                       body.Description = docName;
                       body.MimeType =  "application/xlsx";  //"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";//
                       string newTitle = System.IO.Path.GetFileName("Rack Delivery Plan" + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Year.ToString("00") + ".xlsx");
   
                        // Google.Apis.Drive.v3.Data.File body1 = copyFile(service, fileid, newTitle);
                       //  moveFile(fileId, toFolder);
                       filename = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Rack Delivery Plan Template.xlsx";// "E:\\Ruchi\\b2bdevMaintainGoogleDoc\\MaintainGoogleDoc\\MaintainGoogleDoc\\bin\\Debug\\MIM_Menlo Demand TrackerTemplate.xlsx";
                       filename = filename.Replace("file:\\", " ");
                       // MessageBox.Show("filename=" + filename);
                       System.IO.Stream fileStream = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                       byte[] byteArray = new byte[fileStream.Length];
                       fileStream.Read(byteArray, 0, (int)fileStream.Length);
                       System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
                       FilesResource.UpdateMediaUpload request = service.Files.Update(body, fileId, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
                        //  request.co.Convert = true;
                      request.Upload();
                  //     FilesResource.CreateMediaUpload request = service.Files.Create(body, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                  //     request.Upload();
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
              Cursor.Current = Cursors.Default;
            }



        public void moveFile(string fileId,string toFolder)
        {
          
            // Retrieve the existing parents to remove
            var getRequest = service.Files.Get(fileId);
            getRequest.Fields = "parents";
            var file = getRequest.Execute();
            var previousParents = String.Join(",", file.Parents);

            var getRequest1 = service.Files.Get(fileId);            
        //  var file1 = getRequest1.Execute();
        //  var updateRequest1 = service.Files.Update(new Google.Apis.Drive.v3.Data.File(), fileId);
        //   updateRequest1.Fields = "id, Name";
        //   file1.Name = file1.Name + "111";
        //   file1 = updateRequest1.Execute();
        // Move the file to the new folder

            var CopyRequest = service.Files.Copy(new Google.Apis.Drive.v3.Data.File(), fileId);
            var Copyfile=CopyRequest.Execute();
              
            var updateRequest = service.Files.Update(new Google.Apis.Drive.v3.Data.File(),  Copyfile.Id);
            updateRequest.Fields = "id, parents";
            updateRequest.AddParents = toFolder;
            updateRequest.RemoveParents = previousParents;
            file = updateRequest.Execute();  
           
        }

 /*       private static Google.Apis.Drive.v3.Data.File copyFile(DriveService service, String originFileId, String copyTitle)
        {
            try
            {
                Google.Apis.Drive.v3.Data.File copiedFile = new Google.Apis.Drive.v3.Data.File();
                copiedFile.Name = copyTitle;
                service.Files.Copy(copiedFile, originFileId).Execute();

                var updateRequest = service.Files.Update(new Google.Apis.Drive.v3.Data.File(),Co);
                 updateRequest.AddParents = toFolder;
                
                file = updateRequest.Execute();

                // Rename the file.
                // FilesResource.PatchRequest request = service.Files.Patch(file, fileId);
                // File updatedFile = request.Fetch();

                return service.Files.Copy(copiedFile, originFileId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
                return null;
            }
        }*/

        //public static void RemoveFileFromFolder(DriveService service, String folderId, String fileId)
        //{
        //    try
        //    {
        //        service.Parents.Delete(fileId, folderId);
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine("An error occurred: " + e.Message);
        //    }
        //}

        private void ultraCombo1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {

        }

        private void btnDownloadFile_Click(object sender, EventArgs e)
        {
            //dr = new Google.Apis.Drive.v2.DriveService(CreateAuthenticator());
            //if (dr != null)
            //{
            //    string fileid = "";
            //    List<Google.Apis.Drive.v2.Data.File> fileList = Google.Apis.Util.Utilities.RetrieveAllFiles(dr);
            //    foreach (Google.Apis.Drive.v2.Data.File item in fileList)
            //    {
            //        if (item.Title == "Copy of MIM-Menlo Demand Tracker")//tartsWith(fileName))//"MIM NPI:" + projectName)
            //        {
            //            fileid = item.Id;
            //            //     DataSet ds = getDataSet("select  dbo.udf_GetISOWeekNumberFromDate(getdate())");
            //            //    string last_week_num = ds.Tables[0].Rows[0][0].ToString();
            //            //   File copiedfile = copyFile(dr, fileid, "WeekNum_" + last_week_num + "_" + fileName);

            //            //  renameFile(DriveService service, String fileId, String newTitle)
            //            //  renameFile(dr,fileid,fileName + "_Archive");
            //            //  copyFile(dr, fileid, fileName + "_Archive");
            //            //  insertFileIntoFolder(dr, archiveParentid, copiedfile.Id);
            //            //  MessageBox.Show(copiedfile.UserPermission.AuthKey.ToString());
            //            //   DeleteFile(dr, fileid);
            //            break;
            //        }
            //    }
            //    Google.Apis.Drive.v2.Data.File fs = dr.Files.Get(fileid).Fetch();
            //    string downloadurl = fs.ExportLinks["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"];
            //    //downloadurl= downloadurl.Replace("xlsx", "csv");
            //    var fileStream = new System.IO.FileStream("e:\\ruchi\\DriveFile.csv", System.IO.FileMode.Create, System.IO.FileAccess.Write);
            //    byte[] buffer = new byte[4096];
            //    byte[] result;
            //    System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(new Uri(downloadurl));
            //    CreateAuthenticator().ApplyAuthenticationToRequest(request);
            //    System.Net.WebClient webClient = new System.Net.WebClient();
            //    webClient.DownloadFile(downloadurl, @"e:\\ruchi\\DriveFile1111.xlsx");
            //    //System.Net.HttpWebResponse response = (System.Net.HttpWebResponse)request.GetResponse();
            //    using (System.Net.WebResponse response = request.GetResponse())
            //    {
            //        using (System.IO.Stream responseStream = response.GetResponseStream())
            //        {
            //            using (System.IO.MemoryStream memoryStream = new System.IO.MemoryStream())
            //            {
            //                int count = 0;
            //                do
            //                {
            //                    count = responseStream.Read(buffer, 0, buffer.Length);
            //                    memoryStream.Write(buffer, 0, count);
            //                } while (count != 0);
            //                result = memoryStream.ToArray();
            //                write_data_to_excel(result);
            //            }
            //        }
            //    }
            //    //  response.ContentType = "application/vnd.ms-excel";

            //    //System.IO.Stream stream = response.GetResponseStream();

            //    //  stream.CopyTo(fileStream);

            //    //  byte[] byteArray = new byte[stream.Length];
            //    //   stream.Read(byteArray, 0, (int)stream.Length);
            //    // System.IO.MemoryStream streamMem = new System.IO.MemoryStream(byteArray);
            //    //  fileStream.Write(byteArray, 0, Convert.ToInt32(fileStream.Length));

            //    fileStream.Close();
            //}
        }


        private void write_data_to_excel(byte[] input)
        {
            // System.IO.StreamWriter str = new System.IO.StreamWriter("stockdata.xls");
            System.IO.FileStream f = System.IO.File.OpenWrite("e:\\ruchi\\stockdata.xlsx");
            for (int i = 0; input.Length > i; i++)
            {
                f.Write(input, 0, input.Length);
                //str.WriteLine(input[i].ToString());
            }
            f.Close();
        }


        private void button6_Click(object sender, EventArgs e)
        {
            //dr = new Google.Apis.Drive.v2.DriveService(CreateAuthenticator());
            //if (dr != null)
            //{
            //    List<Google.Apis.Drive.v2.Data.File> fileList = Google.Apis.Util.Utilities.RetrieveAllFiles(dr);
            //    string Parentid = "";
            //    string fileid = "";

            //    foreach (Google.Apis.Drive.v2.Data.File item in fileList)
            //    {
            //        if (item.Title.StartsWith("MIM Menlo Demand Tracker"))//"MIM NPI:" + projectName)
            //        {
            //            Parentid = item.Id;
            //            //renameFile(dr, fileid, promptValue + "_Archive");
            //            break;
            //        }

            //    }
            //    string newTitle = "Rack Delivery Plan" + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Year.ToString("00") + ".xlsx";
            //    foreach (Google.Apis.Drive.v2.Data.File item in fileList)
            //    {

            //        if (item.Title.StartsWith("Rack Delivery Plan07242014.xlsx"))//"MIM NPI:" + projectName)
            //        {
            //            fileid = item.Id;
            //            //renameFile(dr, fileid, promptValue + "_Archive");
            //            break;
            //        }
            //    }
            //    RemoveFileFromFolder(dr, Parentid, fileid);
            //}

        }
    }
}