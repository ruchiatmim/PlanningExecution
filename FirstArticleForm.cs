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
    public partial class FirstArticleForm : Form
    {

         string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
          Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
          Microsoft.Office.Interop.Excel.Workbooks excelWorkbooks = null;
          Microsoft.Office.Interop.Excel.Sheets excelSheets = null;
          Microsoft.Office.Interop.Excel.Worksheet sheet = null;
          Microsoft.Office.Interop.Excel.Application excelApp1 = null;
        string filename1="";

        public FirstArticleForm()
        {
            InitializeComponent();            
        }


        string foldername = "\\\\mvfile\\MV Reports\\MAster Data & forms\\PlanningExecutionForms\\Templates\\";
public void create_file()
{

    System.Diagnostics.Process p = new System.Diagnostics.Process();
    bool createdFile = false;

    try
    {
       excelApp1 = new Microsoft.Office.Interop.Excel.Application();

       string filename = foldername + "First Article Sheet Template.xlsx";

     //   string filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MIM_Menlo Demand TrackerTemplate.xlsx";
        //filename = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Demand Tracker1_"+DateTime.Now.Month.ToString("00") + "_" + DateTime.Now.Day.ToString("00") + "_" + DateTime.Now.Year.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") +".xlsx";
        string workbookPath = filename.Replace("file:", "\\");

        filename1 = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\First Article Sheet Template_" + DateTime.Now.Month.ToString("00") + "_" + DateTime.Now.Day.ToString("00") + "_" + DateTime.Now.Year.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + ".xlsx";
        excelWorkbooks = excelApp1.Workbooks;
        excelWorkbook = excelWorkbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        //    excelWorkbook = excelApp1.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        excelWorkbook.SaveAs(filename1, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
  
        excelSheets = excelWorkbook.Worksheets;
        sheet = excelSheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
        String strQuery = "   SELECT GDOC_FirstArticleSchedule.RefreshDate,GDOC_FirstArticleSchedule.ProdSite, GDOC_FirstArticleSchedule.SFSRouting, GDOC_FirstArticleSchedule.PartMap, GDOC_FirstArticleSchedule.WorkInstr, GDOC_FirstArticleSchedule.QADoc, GDOC_FirstArticleSchedule.Pack, GDOC_FirstArticleSchedule.Tooling, GDOC_FirstArticleSchedule.OSPPO, GDOC_FirstArticleSchedule.Barcode, GDOC_FirstArticleSchedule.Test, GDOC_FirstArticleSchedule.WorkOrdNum, GDOC_FirstArticleSchedule.AsmGPN, GDOC_FirstArticleSchedule.description, GDOC_FirstArticleSchedule.Qty, GDOC_FirstArticleSchedule.SchedAsmDate, GDOC_FirstArticleSchedule.SchedWeek, GDOC_FirstArticleSchedule.DemandSource, GDOC_FirstArticleSchedule.CommCode, GDOC_FirstArticleSchedule.RefGPN, GDOC_FirstArticleSchedule.PlanGroup, GDOC_FirstArticleSchedule.FormFactor, GDOC_FirstArticleSchedule.Platform, GDOC_FirstArticleSchedule.BP_BuildID, GDOC_FirstArticleSchedule.BP_TLAGPN, GDOC_FirstArticleSchedule.BP_POPDate, GDOC_FirstArticleSchedule.BP_FinalDestination, GDOC_FirstArticleSchedule.BP_Cluster, GDOC_FirstArticleSchedule.BP_ProjCode, GDOC_FirstArticleSchedule.BP_Owner   FROM MIMDIST.dbo.GDOC_FirstArticleSchedule GDOC_FirstArticleSchedule  ORDER BY GDOC_FirstArticleSchedule.SchedAsmDate    ";
        DataSet dsFirstArticle = getdataSet(strQuery);
        int col = 0;

        for (int i = 0; i < dsFirstArticle.Tables[0].Rows.Count; i++)
        {
            //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());

            for (int j = 0; j < dsFirstArticle.Tables[0].Columns.Count; j++)
            {
                sheet.Cells[i + 3, j + 1] = dsFirstArticle.Tables[0].Rows[i][j].ToString();
            }
        }
        excelWorkbook.Save();
     //   excelWorkbook.Close(true, false, Type.Missing);
        createdFile = true;

        p.StartInfo.FileName = filename1;
        //p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
        p.Start();              
      //  System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp1);             
             
    }
    catch (Exception e)
    {
       Console.Write(e.Message.ToString());
    }
    finally
    {
        //Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
        /* excelWorkbook.Close(true, false, Type.Missing);
          System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheets);
          System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
          System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbooks);
          System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp1);*/
    }  
 }

      

DriveService service;
string fileid = "";
string Parentid = "";
string Collection = "MIM First Article Folder";
string docName = "First Article Sheet (2015W34)";

public void UploadDoc()
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


            if (file.Name == "MIM First Article Folder")
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
            if (file.Name == "First Article Sheet (2015W34)")
            {
                fileId = file.Id;
                break;

            }
        }
        pageToken = result.NextPageToken;
    } while (pageToken != null);


              try
                   {
                       Google.Apis.Drive.v3.Data.File body = new Google.Apis.Drive.v3.Data.File();
                       body.Name = "TEST CSV DOC";
                       body.Description = "TEST CSV DOC";
                       body.MimeType ="text/csv";                                   

                       filename1 = filename1.Replace("file:\\", " ");
                       filename1 = "e:\\ACCPurchaseOrderDetail.csv";
                       // MessageBox.Show("filename=" + filename);
                       System.IO.Stream fileStream = System.IO.File.Open(filename1, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                       byte[] byteArray = new byte[fileStream.Length];
                       fileStream.Read(byteArray, 0, (int)fileStream.Length);
                       System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
                      // FilesResource.UpdateMediaUpload request = service.Files.Update(body, fileId, stream, "text/csv");//application/x-vnd.oasis.opendocument.spreadsheet");
                       FilesResource.CreateMediaUpload request1 = service.Files.Create(body, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
                    
                       request1.Upload();
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

        private void button1_Click(object sender, EventArgs e)
        {
            create_file();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            UploadDoc();
        }

    }
}
