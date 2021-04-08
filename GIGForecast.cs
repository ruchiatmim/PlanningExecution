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

using Google.Apis.Util;
using System.Net;

namespace Version3
{
    public partial class GIGForecast : Form
    {

        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        public GIGForecast()
        {
            InitializeComponent();
            getauthorization();
        }

       // DriveService service;
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
                string title = "GIG-MIM build plan";
                request.Q = "fullText contains '" + title + "'";
               // string parent_id = "0B1auUUiGcbbBOEIyNjg0alkwc3c";
               // request.Q = "'" + parent_id + "' in parents";
                request.Spaces = "Drive";
                request.Fields = "nextPageToken, files(id, name)";
                request.PageToken = pageToken;

                var result = request.Execute();
                foreach (var file in result.Files)
                {
                    if (file.Name.StartsWith("GIG-MIM build plan"))
                    {
                        ListViewItem item = new ListViewItem(new string[2] { file.Name, file.Id });
                        googleshareDoc = file.Name;
                        this.spreadsheetListView1.Items.Add(item);

                    }
                }
                pageToken = result.NextPageToken;
            } while (pageToken != null);

      /*      string fileName = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\currentRackDeliveryPlan.xlsx";
            fileName = fileName.Replace("file:\\", "");

            Google.Apis.Drive.v3.Data.File file1 = service.Files.Get(fileId).Execute();
            if (this.downloadfile(service, fileId, fileName))
            {

                getallSheet(fileName);
                //  MessageBox.Show("File has been downloaded");

            }*/

        }





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


   /*     public void getallSheet(string FileName)
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
        */
        DriveService service;
        string fileid = "";
        string Parentid = "";
        string Collection = "MIM Menlo Demand Tracker";
        string docName = "Net Supply Plan";
        string googleshareDoc = "";
/*
        public void ImportsharedocToExcel()
        {

            UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
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

            FilesResource.ListRequest lstreq = service.Files.List();
            do
            {
                try
                {
                    lstreq.Q = " Title contains 'GIG-MIM build plan'";
                    FileList fileList = lstreq.Execute();
                    foreach (Google.Apis.Drive.v2.Data.File fileitem in fileList.Items)
                    {
                        if (fileitem.Title.StartsWith("GIG-MIM build plan"))
                        {
                            ListViewItem item = new ListViewItem(new string[2] { fileitem.Title, fileitem.Id });
                            googleshareDoc = fileitem.Title;
                            this.spreadsheetListView1.Items.Add(item);
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
        }
        */
        string saveTo = "";

/*
        public Boolean downloadFile(DriveService _service, File _fileResource, string _saveTo)
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
                    MessageBox.Show("An error occurred: " + e.Message);
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
        string table_name = "";
        string ALLColumns = "";
        public void showInGrid()
        {
            System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + saveTo + "';Extended Properties=Excel 12.0;");
            System.Data.OleDb.OleDbDataAdapter MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Forecast$]", MyConnection);
            MyCommand.TableMappings.Add("Table", "TestTable");
            System.Data.DataSet DtSet = new System.Data.DataSet();
            MyCommand.Fill(DtSet);
            DataTable dtShareColl = DtSet.Tables[0];
            MyConnection.Close();
            for (int i = 0; i < dgvNewforecast.ColumnCount; i++)
                dgvNewforecast.Columns[i].Width = 200;

            getColumnName(dtShareColl);
        }

    
        string filename;
        private void spreadsheetListView1_SelectedIndexChanged(object sender, EventArgs e)
        {          

        }
        
        string columnName;
        string sharecolumnName;
        string[] colDatatype = new string[100];
        string[] colName = new string[100];
        private void getColumnName(DataTable dtShareColl)
        {

           
            System.Data.DataSet ds = new System.Data.DataSet();
            string sqlQry = "select  *  from [m_gig_header_mapping] select part_no from inv_master ";
            try
            {
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlCommand cmd = new SqlCommand(sqlQry, cn);
                cmd.CommandTimeout = 0;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                ds = new System.Data.DataSet();
                da.Fill(ds);
                cn.Close();
                //MessageBox.Show("File has been uploaded successfully");
            }
            catch (Exception ex)
            {
                //MessageBox.Show("There is some error to save. Please check." + ex.Message.ToString());
            }

            DataTable dtGrid=new DataTable();
            int j = 0;
            foreach (DataRow dr in ds.Tables[0].Rows)
            {                
                columnName = columnName + dr["table_field_name"] + ",";
                sharecolumnName = sharecolumnName + dr[3] + ",";
                dtGrid.Columns.Add(dr[3].ToString().ToUpper());
                colDatatype[j] = dr[4].ToString().ToLower();
                colName[j] = dr[3].ToString().ToLower();
                j = j + 1;
            }

            dtGrid.Columns.Add("NOT EXIST");

            columnName = columnName.Substring(0, columnName.Length - 1);
            foreach (DataRow sourcerow in dtShareColl.Rows)
            {
                DataRow destRow = dtGrid.NewRow();
                for (int k = 0; k < j; k++)
                {
                    try
                    {
                        if (colDatatype[k] == "integer" && (sourcerow[colName[k]].ToString() == "" || sourcerow[colName[k]] == null))
                            sourcerow[colName[k]] = 0;
                        else if (colDatatype[k] == "varchar" && (sourcerow[colName[k]].ToString().Trim() == "" || sourcerow[colName[k]] == null || sourcerow[colName[k]].ToString().Trim() == "0"))
                            sourcerow[colName[k]] = "";
                        if (colName[k].Contains("date"))
                            destRow[colName[k]] = Convert.ToDateTime(sourcerow[colName[k]]).ToShortDateString();
                        else if (colName[k].Contains("item number") && sourcerow[colName[k]].ToString().StartsWith("7"))
                            destRow[colName[k]] = "0"+ sourcerow[colName[k]];
                        else
                            destRow[colName[k]] = sourcerow[colName[k]];
                       
                    }
                    catch (Exception ex)
                    {

                    }                

                }
                if (destRow[colName[0]].ToString() == "" && destRow[colName[1]].ToString() == "" && destRow[colName[2]].ToString() == "")
                    break;
                DataRow[] drExist = ds.Tables[1].Select("part_no='" + destRow["Item Number"].ToString() + "'");
          /*      if (destRow["Item Number"].ToString() != destRow["TLA GPN"].ToString())
                {
                    if (drExist.Length < 2)
                    {
                        destRow["NOT EXIST"] = "Y";
                        btnUploadNewForecast.Enabled = false;
                    }
                    else
                        destRow["NOT EXIST"] = "N";
                }
                else
                {*/
                    if (drExist.Length < 1)
                    {
                        destRow["NOT EXIST"] = "Y";
                        string part = destRow[colName[2]].ToString().Trim();
                        if (part != "")
                        btnUploadNewForecast.Enabled = false;
                    }
                    else
                        destRow["NOT EXIST"] = "N";
              //  }
                dtGrid.Rows.Add(destRow);
            }           

            dgvNewforecast.DataSource = dtGrid;            
           // string tablename = "m_nsp_plan";
            // getExcelcolumnNameNSP(tablename);
            lblNumOfRec.Text = "Number of Records :" + dtGrid.Rows.Count.ToString();            
        }




        private void button9_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (spreadsheetListView1.SelectedItems.Count > 0)
            {

                filename = spreadsheetListView1.SelectedItems[0].SubItems[1].Text;
                //  googleshareDoc = spreadsheetListView.;
                //  fileId = spreadsheetListView.SelectedItems[0]..Text;
              //  Google.Apis.Drive.v3.Data.File file1 = service.Files.Get(filename).Execute();
                saveTo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\currentForecast" + System.DateTime.Now.Day.ToString("00") + System.DateTime.Now.Month.ToString("00") + System.DateTime.Now.Year.ToString("00") + System.DateTime.Now.Hour.ToString("00") + System.DateTime.Now.Minute.ToString("00") + System.DateTime.Now.Second.ToString("00") + ".xlsx";
                saveTo = saveTo.Replace("file:\\", "");
                if (downloadfile(this.service, filename, saveTo))
                {
                    showInGrid();
                }

            }
            Cursor.Current = Cursors.Default;
        }

        private void btnUploadNewForecast_Click(object sender, EventArgs e)
        {

            Cursor.Current = Cursors.WaitCursor;
            DataTable dtBuild = this.dgvNewforecast.DataSource as DataTable;
            string sqlDelete = "  delete from m_gig_forecast_table ";
            string sqlSave = "";
            string sqlSaveFixed = "";
            string[] colName = sharecolumnName.Split(',');
          //  string[] datetype = dbDateType.Split(',');
           DataRow[] drSelect= dtBuild.Select("[ITEM NUMBER] <>''");
           for (int i = 0; i < drSelect.Length; i++)
            {
                if (drSelect[i][0].ToString().Trim() != "")
                {
                    sqlSaveFixed = sqlSaveFixed + "  insert into m_gig_forecast_table  ( " + columnName + /* forecast_tab_name,region,category,planning_item,delivery_date,qty*/ " )  values (";
                    for (int j = 0; j < colName.Length-1; j++)
                    {
                        if (j!=0)
                        sqlSaveFixed = sqlSaveFixed + ",";
                        if (colDatatype[j] == "int")
                        //   if (j == colname.Length - 1)
                        {
                            //sqlSaveFixed = sqlSaveFixed + "'" + dtBuild.Rows[i][colname[j]].ToString().Replace(",", "") + "'";
                            int qty = 0;
                            decimal decQty = Convert.ToDecimal(drSelect[i][colName[j]].ToString().Replace(",", ""));
                            qty = Decimal.ToInt32(decQty);
                            sqlSaveFixed = sqlSaveFixed + "'" + qty + "'";
                        }
                        else
                        {
                            sqlSaveFixed = sqlSaveFixed + "'" + drSelect[i][colName[j]].ToString() + "'";
                        }
                    }

                    sqlSaveFixed = sqlSaveFixed + ")";
                }
            }

            sqlSave = sqlDelete + sqlSaveFixed;
            string sqlupdate = " update MRP_Process_step_table set m_mrp_Plan_update_time=GETDATE()  where m_mrp_plan_name='GIG Forecast' ";
           
                sqlSave = sqlSave + sqlupdate;
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlTransaction tran = cn.BeginTransaction(); ;
            try
            {
                
                SqlCommand cmd = new SqlCommand(sqlSave, cn);
                cmd.Transaction = tran;
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();
                tran.Commit();
                cn.Close();
                MessageBox.Show("File has been uploaded successfully");
            }
            catch (Exception ex)
            {
                tran.Rollback();
                MessageBox.Show("There is some error to save. Please check." + ex.Message.ToString());
            }
            Cursor.Current = Cursors.Default;
        }

        private void dgvNewforecast_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            

            if ((e.ColumnIndex == dgvNewforecast.Columns.Count - 1) && (e.Value == "Y"))
            {

                for (int col = 0; col < dgvNewforecast.Columns.Count; col++)
                {
                   
                    dgvNewforecast.Rows[e.RowIndex].Cells[col].Style.BackColor = Color.Silver;
                }

            }
            else
                dgvNewforecast.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
           
        }
      
    }
}
