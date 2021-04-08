using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Infragistics.Win;





using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Download;

namespace Version3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            getauthorization();
        }

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
            var service = new DriveService(new BaseClientService.Initializer()
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
                request.Q ="'"+parent_id+"' in parents";
                request.Spaces = "Drive";
                request.Fields = "nextPageToken, files(id, name)";
                request.PageToken = pageToken;

                var result = request.Execute();
                foreach (var file in result.Files)
                {
                    if (file.Name == "Rack Delivery Plan [go/rackdeliveryplan]")
                    {
                        fileId = file.Id;
                        Google.Apis.Drive.v3.Data.File fileG = service.Files.Get(fileId).Execute();
                                             
                       
                        break;
                      
                    }
                }
                pageToken = result.NextPageToken;
            } while (pageToken != null);

            string fileName = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Templates\\currentRackDeliveryPlan.xls";
            fileName = fileName.Replace("file:\\", "");
           if  (downloadfile(service, fileId, fileName))
           {

               MessageBox.Show("File has been downloaded");

           }

        }


/*
         
                        // Define parameters of request.
                        FilesResource.ListRequest listRequest = service.Files.List();
                        listRequest.PageSize = 1000;
                        listRequest.Fields = "nextPageToken, files(id, name)";
            listRequest.mi
                        listRequest.Q = "full text contain  'MIM Menlo Demand Tracker'";


                        // List files.
                        IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                            .Files;
                        Console.WriteLine("Files:");
                        if (files != null && files.Count > 0)
                        {
                              foreach (var file in files)
                            {string FileId=file.Id;

                            if (file.Id == "1-w9r2-8mtTVu-iwfEOS6rc9JVhI_LUfUS7pKtyFF-2M")
                                {                        
                                    // Retrieve the existing parents to remove
                                    var getRequest = service.Files.Get(FileId);
                                    getRequest.Fields = "parents";
                                  var v_files = getRequest.Execute();
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("No files found.");
                        }
                        Console.Read();

                    }
            */


        public bool downloadfile(DriveService service, string fileId, string fileName)
        {

            var request = service.Files.Export(fileId, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            var streamvar = new System.IO.MemoryStream();
            // Add a handler which will be notified on progress changes.
            // It will notify on each chunk download and when the
            // download is completed or failed.
            string strprogress = "";
            request.MediaDownloader.ProgressChanged += (IDownloadProgress progress) =>
                    {
                        switch (progress.Status)
                        {
                            case DownloadStatus.Downloading:
                                {
                                    strprogress = "Downloaded Bytes " + progress.BytesDownloaded;
                                    break;
                                }
                            case DownloadStatus.Completed:
                                {
                                    strprogress = "Download complete.";
                                    break;
                                }
                            case DownloadStatus.Failed:
                                {
                                    strprogress = "Download failed.";
                                    break;
                                }
                        }
                    };
            Stream str = request.ExecuteAsStream();//.Download(streamvar);
           
           

           // var buffer = System.IO.File.ReadAllBytes("e:\\ShipTracker.xlsx");


 //streamvar = new MemoryStream(buffer);
           
          
                   using (Stream file = System.IO.File.Create("e:\\fileNamesuccess.xlsx"))
{
    byte[] buffer = new byte[8 * 1024];
    int len;
    while ((len = str.Read(buffer, 0, buffer.Length)) > 0)
    {
        file.Write(buffer, 0, len);
    }    
                 return   true;
}
               
               
             
         
            return false;
        }
        int sheetnum=0;
        Microsoft.Office.Interop.Excel.Workbook excelWorkbook;
        Microsoft.Office.Interop.Excel.Sheets excelSheets;
        public void getallSheet(string FileName)
        {
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
/*
        private void button6_Click(object sender, System.EventArgs e)
        {

        }

        private void btnDownloadFile_Click(object sender, System.EventArgs e)
        {

        }

        private void button1_Click(object sender, System.EventArgs e)
        {

        }

        private void btnUpdateTable_Click(object sender, System.EventArgs e)
        {

        }

      
        private void btnGetDataFromDoc_Click_1(object sender, System.EventArgs e)
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

        public void showInGrid(string table_name)
        {
            //  table_name = "Rack Delivery Plan";
            //   saveTo="E:\\Ruchi\\NewPlanningExceution\\MaintainGoogleDoc\\MaintainGoogleDoc\\bin\\Debug\\templateRack Delivery Plan Template.xlsx";
            System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + saveTo + "';Extended Properties=Excel 12.0;");
            System.Data.OleDb.OleDbDataAdapter MyCommand = new System.Data.OleDb.OleDbDataAdapter("select  * from [" + table_name + "$]", MyConnection);
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
                            if (drShare[drMapping["doc_header_alpha"].ToString()].ToString().Trim() != "")
                            {
                                try
                                {
                                    drGrid[drMapping["field_name"].ToString()] = drShare[drMapping["doc_header_alpha"].ToString()].ToString();
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
                    band.Columns[dc.ColumnName].MaskClipMode =UltraWinMaskedEdit.MaskMode.IncludeLiterals;
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

        private void btnUpdateTable_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            ListItem1 lst = lstWorksheet.SelectedItem as ListItem1;
            DataTable dt = dgDownload.DataSource as DataTable;
            string sqlval = "    ";

            string sqlIns = "  exec  sp_mim_menlo_demand_tracker ";//demand_source,   build_id, tla_part_num, tla_qty, del_product, del_platform, del_part_num, del_qty, menlo_fill_qty, menlo_status,final_destination,plan_model,order_number,comments) values ";
            //build_id,rack_part_num,menlo_battery_12pack, menlo_battery_24pack,menlo_battery_juicebox,menlo_battery_PSU_boat, menlo_configured_rack,population_complete,mim_notes,menlo_notes) values ";
            // string sqlUpd=" update [m_mim_menlo_demand_tracker_temp] set tab_name='" + Machine Rack Plan + "',"

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

*/
    }
}