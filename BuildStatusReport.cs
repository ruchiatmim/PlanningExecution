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



namespace Version3
{
    public partial class BuildStatusReport : Form
    {
        public BuildStatusReport()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Process.Start("\\\\192.168.0.29\\c$\\SharedocJob\\BuildIDStatusReport\\bin\\Debug\\BuildIDStatusReport.exe");
            create_file();
            uploadFile();
        }

        string filename1 = "";
        string filename = "";
        int sheetnum = 0;
        string tempfilename = "\\\\mvfile\\MV Reports\\MAster Data & forms\\PlanningExecutionForms\\Templates\\Build ID 4 Week Summary template.xlsx";
        string foldername= "\\\\mvfile\\MV Reports\\MAster Data & forms\\PlanningExecutionForms\\Templates\\";

           Microsoft.Office.Interop.Excel.Worksheet sheet;
            Microsoft.Office.Interop.Excel.Sheets excelSheets;

          
        public void create_file()
            {
                Microsoft.Office.Interop.Excel.Application excelApp1 = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp1.Workbooks.Add();

      


          System.Diagnostics.Process p = new System.Diagnostics.Process();
             try
             {
                 excelApp1 = new Microsoft.Office.Interop.Excel.Application();
                 //string workbookPath = "E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\VendorSheet.xls"; Microsoft.Office.Interop.Excel.XlPlatform.xlWindows  
                 string filename =  tempfilename;//Menlo_Metrics_Kit_SnapshotTemplate.xlsx";
                 string workbookPath = filename.Replace("file:\\", "");
             
                 filename1 = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\" +workbookPath.Replace(foldername, "").Replace(".xlsx", "") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Year.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + ".xlsx";
                 //  excelWorkbook = excelApp1.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                 excelWorkbook = excelApp1.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                 excelWorkbook.SaveAs(filename1, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                 sheetnum = 0;
                 excelSheets = excelWorkbook.Worksheets;
                 sheetnum = sheetnum + 1;

                 sheet = excelSheets[sheetnum] as Microsoft.Office.Interop.Excel.Worksheet;
                // updateSheetEx(sheet);
                 updateSheet(sheet);
                 excelWorkbook.Save();

             }

             catch (Exception e)
             {
                 // MessageBox.Show(e.Message.ToString());
             }
            finally
            {
                //Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
                excelWorkbook.Close(true, false, Type.Missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp1);
            }
        }

        private void updateSheetEx(Microsoft.Office.Interop.Excel.Worksheet sheet)
        {
            String strQuery = "SELECT     TOP (100) PERCENT RefreshDate, BuildID, POPDate, BOMStatus, BuildDate, TrayLocked, RackLocked, TLAGPN, TLAQty, Cluster, Destination, Owner, GPNDesc FROM         dbo.excel_4WK_buildID_summary AS excel_4WK_buildID_summary   ORDER BY BuildDate, BuildID ";
            string cnt = "data source=mverp;initial catalog=mimdist;user id=sa;password=mimi~100;pooling=true;max pool size=100;min pool size=1;";
            Microsoft.Office.Interop.Excel.QueryTable oQryTable = sheet.QueryTables.Add("OLEDB;Provider=sqloledb;" + cnt, sheet.Range["A1"], strQuery);
            oQryTable.RefreshStyle = Microsoft.Office.Interop.Excel.XlCellInsertionMode.xlInsertEntireRows; // 2; //' xlInsertEntireRows = 2
            oQryTable.Refresh(false);
         

        }

        public void updateSheet(Microsoft.Office.Interop.Excel.Worksheet sheet)
        {
            String strQuery = "SELECT     TOP (100) PERCENT RefreshDate, BuildID, POPDate, BOMStatus, BuildDate, TrayLocked, RackLocked, TLAGPN, TLAQty, Cluster, Destination, Owner, GPNDesc FROM         dbo.excel_4WK_buildID_summary AS excel_4WK_buildID_summary   ORDER BY BuildDate, BuildID ";

            System.Data.DataSet dsBPDetail = getDataSet(strQuery);
            int col = 0;

            for (int i = 0; i < dsBPDetail.Tables[0].Rows.Count; i++)
            {
                //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());

                for (int j = 0; j < dsBPDetail.Tables[0].Columns.Count; j++)
                {
                    sheet.Cells[i + 2, j + 1] = dsBPDetail.Tables[0].Rows[i][j].ToString();
                }
            }

        }


        string constr = "Data Source=mverp;Initial Catalog=MIMDIST;user id=sa;password=mimi~100;";
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
             DriveService service;
        public void uploadFile()
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


                    if (file.Name == "MIM Build ID Rolling 4WK Summary")
                    {
                        MainParentid = file.Id;
                        break;
                    }
                }



                pageToken = result.NextPageToken;
            } while (pageToken != null);



            do
            {
                var request = service.Files.List();
                // request.Q = "mimeType = 'application/vnd.google-apps.folder' ";
                string parent_id = MainParentid;
                request.Q = "'" + parent_id + "' in parents";
                request.Spaces = "Drive";
                request.Fields = "nextPageToken, files(id, name)";
                request.PageToken = pageToken;

                var result = request.Execute();
                foreach (var file in result.Files)
                {
                    if (file.Name == "Build ID 4 Week Summary")
                    {
                        fileId = file.Id;
                        break;

                    }
                }
                pageToken = result.NextPageToken;
            } while (pageToken != null);

            Google.Apis.Drive.v3.Data.File body = service.Files.Get(fileId).Execute();        


            try
            { 
                filename1 = filename1.Replace("file:\\", " ");
                // MessageBox.Show("filename=" + filename);
                System.IO.Stream fileStream = System.IO.File.Open(filename1, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                byte[] byteArray = new byte[fileStream.Length];
                fileStream.Read(byteArray, 0, (int)fileStream.Length);
                System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
                FilesResource.UpdateMediaUpload request = service.Files.Update(body, fileId, stream, "application/x-vnd.oasis.opendocument.spreadsheet");

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

        
            
          /*   
             Google.Apis.Drive.v2.Data.File body = service.Files.Get(fileid).Execute();         
                 
                  try
                {               
                  
                  

                  //  filename = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\" + "\\MIM_Menlo Demand TrackerTemplate.xlsx";// System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\MIM_Menlo Demand TrackerTemplate.xlsx";// "E:\\Ruchi\\b2bdevMaintainGoogleDoc\\MaintainGoogleDoc\\MaintainGoogleDoc\\bin\\Debug\\MIM_Menlo Demand TrackerTemplate.xlsx";
                    filename = filename1.Replace("file:\\", " ");
                  //  filename1 = filename1.Replace("file:\\", " ");
                    // MessageBox.Show("filename=" + filename);
                    System.IO.Stream fileStream = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                    byte[] byteArray = new byte[fileStream.Length];
                    fileStream.Read(byteArray, 0, (int)fileStream.Length);
                    System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
                    FilesResource.UpdateMediaUpload request = service.Files.Update(body, fileid, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
                    //   FilesResource.InsertMediaUpload request = dr.Files.Insert(body, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
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

                }*/
            }
        


  /*     
            public  string  getchildInFolder(DriveService service,   String folderId)
        {
            ChildrenResource.ListRequest request = service.Children.List(folderId);
            string file_id = "";
            do
            {
                try
                {
                    ChildList children = request.Execute();

                    foreach (ChildReference child in children.Items)
                    {
                        Google.Apis.Drive.v2.Data.File file1 = service.Files.Get(child.Id).Execute();
                        if (file1.Title.StartsWith("Build ID 4 Week Summary"))
                        {
                            file_id = child.Id;
                            break;
                          
                        }
                           
                    }
                    request.PageToken = children.NextPageToken;
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                    request.PageToken = null;
                }
            } while (!String.IsNullOrEmpty(request.PageToken));

            return file_id;
        }


        public static ParentReference insertFileIntoFolder(DriveService service, String folderId, String fileId)
        {
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
        

    }
}