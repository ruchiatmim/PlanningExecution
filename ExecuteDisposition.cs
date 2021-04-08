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

using Google.Apis.Util;
using System.Net;
using Microsoft.SqlServer.Dts.Runtime;

namespace Version3
{
    public partial class ExecuteDisposition : Form
    {
        public ExecuteDisposition()
        {
            InitializeComponent();
        }

        private void ultraButton1_Click(object sender, EventArgs e)
        {
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
                credPath = Path.Combine(credPath, ".credentials111/drive-dotnet-quickstart.json");

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

      /*      string pageToken = null;
            string fileId = "";
            do
            {
                var request = service.Files.List();
                // request.Q = "mimeType = 'application/vnd.google-apps.folder' ";
               // string parent_id = "0B1auUUiGcbbBOEIyNjg0alkwc3c";
                request.Q = "name contains 'Disposition Table for MIM 10/07/19' ";
                request.Spaces = "Drive";
                request.Fields = "nextPageToken, files(id, name)";
                request.PageToken = pageToken;

                var result = request.Execute();
                foreach (var file in result.Files)
                {
                   // if (file.Name == "Rack Delivery Plan [go/rackdeliveryplan]")
                   // {
                        fileId = file.Id;
                        break;
                   // }
                }
                pageToken = result.NextPageToken;
            } while (pageToken != null);
            */
          //  string message = fileId;
         string fileId = "1ZnnRVMvx0Ma4GUih2wq3hcJK1vg0LbxoXmKCzwlz1j4";//1Y4hdHIkkVv5SmGtj8QWM-jAccZdEOhhV6Ofx_kb3IBs"; //"1OSr1IqrUjqo3RdAn1WwoNpTf2v64jzY3rKgJ8SS3hf4";//"11emBYnY-_tLQ43TFEAWIsMEtVgwyXEaAmVCPxcnp6Ag";// "1Y4hdHIkkVv5SmGtj8QWM-jAccZdEOhhV6Ofx_kb3IBs";// 
         string fileName="\\\\mimsftp\\e$\\sftp\\msgGoogle\\MDS\\PROD\\inbound\\DISPOSITION\\INBOUND\\Decom Disposition Table For MIM.xlsx";

         if   (downloadFile(service, fileId, fileName))
        {
             runJob();
            //  RunPackage();

          }
 
        }

        public bool downloadFile(DriveService service, string fileId, string fileName)
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

        string constr = ConfigurationManager.ConnectionStrings["MVReportingConnectionString"].ConnectionString;
        private void runJob()
        {

            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlTransaction tran = cn.BeginTransaction();
            try
            {
                SqlCommand cmd = new SqlCommand(" EXEC msdb.dbo.sp_start_job  'Process Disposition' ", cn);
                cmd.CommandTimeout = 0;
                cmd.Transaction = tran;
                cmd.ExecuteNonQuery();                       
                tran.Commit();
                MessageBox.Show("Job has been statrted", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception to upload file");
            }
            cn.Close();
            cn.Dispose();
        }

        private void RunPackage()
        {
            string pkgLocation;
            Package pkg;
            Microsoft.SqlServer.Dts.Runtime.Application app;
            DTSExecResult pkgResults;

            pkgLocation = @"\\mvapps\F$\Integration\SupplierIntegrationSFTP\SupplierIntegrationSFTP\Disposition.dtsx";             
            app = new Microsoft.SqlServer.Dts.Runtime.Application();
            pkg = app.LoadPackage(pkgLocation, null);
            pkgResults = pkg.Execute();

           if (pkgResults == DTSExecResult.Success)
               MessageBox.Show("Package ran successfully");
            else
               MessageBox.Show("Package failed");
       }


    }
}
