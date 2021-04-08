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
    public partial class BOLReport : Form
    {

        string constr = ConfigurationManager.ConnectionStrings["MVIntConnectionString"].ConnectionString;
          Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
          Microsoft.Office.Interop.Excel.Workbooks excelWorkbooks = null;
          Microsoft.Office.Interop.Excel.Sheets excelSheets = null;
          Microsoft.Office.Interop.Excel.Worksheet sheet = null;
          Microsoft.Office.Interop.Excel.Application excelApp1 = null;
        string filename1="";

        public BOLReport()
        {
            InitializeComponent();

            if (create_file())
                UploadDoc();

            
        }


        string foldername = "\\\\mvfile\\MV Reports\\MAster Data & forms\\PlanningExecutionForms\\Templates\\";
public bool create_file()
{

    System.Diagnostics.Process p = new System.Diagnostics.Process();
    bool createdFile = false;
    try
    {
       excelApp1 = new Microsoft.Office.Interop.Excel.Application();

       string filename = foldername + "BOLReportData.xlsx";

     //   string filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MIM_Menlo Demand TrackerTemplate.xlsx";
        //filename = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Demand Tracker1_"+DateTime.Now.Month.ToString("00") + "_" + DateTime.Now.Day.ToString("00") + "_" + DateTime.Now.Year.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") +".xlsx";
        string workbookPath = filename.Replace("file:", "\\");

        filename1 = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\BOL Data _" + DateTime.Now.Month.ToString("00") + "_" + DateTime.Now.Day.ToString("00") + "_" + DateTime.Now.Year.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + ".xlsx";
        excelWorkbooks = excelApp1.Workbooks;
        excelWorkbook = excelWorkbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        //    excelWorkbook = excelApp1.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        excelWorkbook.SaveAs(filename1, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
  
        excelSheets = excelWorkbook.Worksheets;
        sheet = excelSheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
        String strQuery = "   SELECT  [ShipDate]       ,[BOL#]      ,[TruckId]      ,[Seal#]      ,[PalletId]      ,[Container#]      ,[Order#]      ,[GPN]      ,[Qty]      ,[OwnerCode]      ,[ProjectCode]  FROM [PRODDIST].[dbo].[XPO_BOL]  where  [ShipDate] >getdate()-2  order by [ShipDate]  ";
        DataSet dsFirstArticle = getdataSet(strQuery);
        int col = 0;

        for (int i = 0; i < dsFirstArticle.Tables[0].Rows.Count; i++)
        {
            //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());

            for (int j = 0; j < dsFirstArticle.Tables[0].Columns.Count; j++)
            {
                sheet.Cells[i + 2, j + 1] = dsFirstArticle.Tables[0].Rows[i][j].ToString();
            }
        }
        excelWorkbook.Save();
        excelWorkbook.Close(true, false, Type.Missing);   
            
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
    return true;
 }

      

DriveService service;
string fileid = "";
string Parentid = "";
string Collection = "XPO_BOL";
string docName = "BOL Report";

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

 
    string MainParentid = "";

    string pageToken = null;
    string fileId = "";
  
        fileId = "1odr05L9L1bAeu5hZCVP_eH4ie8iH4dm3KmRF3moZdmo";
        MainParentid = "0B1auUUiGcbbBUFZ3d1lkcDF6bVE";
              try
                   {
                       Google.Apis.Drive.v3.Data.File body = new Google.Apis.Drive.v3.Data.File();
                       body.Name = docName;
                       body.Description = docName;
                       body.MimeType = "application/vnd.google-apps.spreadsheet";
                    //   body.Parents = new List<string> { MainParentid };                                   

                       filename1 = filename1.Replace("file:\\", " ");
                       // MessageBox.Show("filename=" + filename);
                       System.IO.Stream fileStream = System.IO.File.Open(filename1, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                       byte[] byteArray = new byte[fileStream.Length];
                       fileStream.Read(byteArray, 0, (int)fileStream.Length);
                       System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);

                       FilesResource.UpdateMediaUpload request = service.Files.Update(body, fileId, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
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
