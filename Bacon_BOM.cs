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
using Microsoft.SqlServer.Dts.Runtime;


namespace Version3
{
    public partial class Bacon_BOM : Form
    {
        public Bacon_BOM()
        {
            InitializeComponent();
        }
        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        private void btnFilter_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
             string sql ="";
          //  if (radBacon.Checked)
         //   {
                sql = " exec [MIMDIST].[dbo].[fs_indent_bld_Bacon_BOM] '" + txtTLA.Text.Trim() + "'  SELECT [SubAsmPN],[Commodity] ,[CompPN] ,[Manufacturer],[MFGPartNum],[QtyPer],[QtyExt],[UnitCost],[ExtCost],[SupplyType],[MFG_LeadTime],[GCMM],[PartDesc],[Head_Message]  FROM [MIMDIST].[dbo].[BaconBOM_for_Interface] order by [SortOrder1] ,[SortOrder2]  ,[SortOrder3]  ,[SortOrder4]    ";

                sql = sql + "  select   distinct [Head_TLAPN] ,[Head_TLADesc] ,[Head_PlanGroup] ,[Head_FormFactor] ,[Head_Platform] ,[Head_AsmStage]   FROM [MIMDIST].[dbo].[BaconBOM_for_Interface] ";
                bindData(sql);                  
                          
                ultraGridBOM.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
              //  ultraGridBOM.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["SubAsmPN"].Width = 230;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["Commodity"].Width = 150;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["CompPN"].Width = 100;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["Manufacturer"].Width = 150;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["MFGPartNum"].Width = 120;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["QtyPer"].Width = 70;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["QtyExt"].Width = 70;
            
                ultraGridBOM.DisplayLayout.Bands[0].Columns["UnitCost"].Width = 70;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["ExtCost"].Width = 70;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["SupplyType"].Width = 100;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["GCMM"].Width = 30;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["PartDesc"].Width = 450;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["Head_Message"].Width = 150;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["MFG_LeadTime"].Width = 70;

                ultraGridBOM.DisplayLayout.Bands[0].Columns["QtyExt"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["QtyPer"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["UnitCost"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["ExtCost"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["SupplyType"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
                ultraGridBOM.DisplayLayout.Bands[0].Columns["MFG_LeadTime"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;                
             
            //}
            //else
            //{
            //    sql = "    select  Top 1 [SharedDocFolderName],TLA_PartNum,TLADesc  FROM [MIMDIST].[dbo].[BaconBOM_for_SharedDoc] where [SharedDocFolderName] is not null SELECT TLA_PartNum,TLADesc,  AssemblyType ,AssyGPN,CommodityCode ,SupplyType,CompGPN ,LeadTime ,BOMQty ,UnitCost,ExtCost,CompDesc    FROM [MIMDIST].[dbo].[BaconBOM_for_SharedDoc] order by sortorder,COMPGPN ";

            //    bindData(sql);

            //    ultraGridBOM.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;

            

         
            //    Infragistics.Win.UltraWinGrid.UltraGridColumn c3 = ultraGridBOM.DisplayLayout.Bands[0].Columns["AssemblyType"];
            //    c3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

            //    Infragistics.Win.UltraWinGrid.UltraGridColumn c4 = ultraGridBOM.DisplayLayout.Bands[0].Columns["AssyGPN"];
            //    c4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

            //    Infragistics.Win.UltraWinGrid.UltraGridColumn c5 = ultraGridBOM.DisplayLayout.Bands[0].Columns["CommodityCode"];
            //    c5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

            //    Infragistics.Win.UltraWinGrid.UltraGridColumn c6 = ultraGridBOM.DisplayLayout.Bands[0].Columns["SupplyType"];
            //    c6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

            //    Infragistics.Win.UltraWinGrid.UltraGridColumn c7 = ultraGridBOM.DisplayLayout.Bands[0].Columns["CompGPN"];
            //    c7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

            //    Infragistics.Win.UltraWinGrid.UltraGridColumn c8 = ultraGridBOM.DisplayLayout.Bands[0].Columns["LeadTime"];
            //    c8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
             
            //    Infragistics.Win.UltraWinGrid.UltraGridColumn c9 = ultraGridBOM.DisplayLayout.Bands[0].Columns["BOMQty"];
            //    c9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

            //    Infragistics.Win.UltraWinGrid.UltraGridColumn c10 = ultraGridBOM.DisplayLayout.Bands[0].Columns["UnitCost"];
            //    c10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

            //    Infragistics.Win.UltraWinGrid.UltraGridColumn c11 = ultraGridBOM.DisplayLayout.Bands[0].Columns["ExtCost"];
            //    c11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

            //    Infragistics.Win.UltraWinGrid.UltraGridColumn c12 = ultraGridBOM.DisplayLayout.Bands[0].Columns["CompDesc"];
            //    c12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

            //         ultraGridBOM.DisplayLayout.Bands[0].Columns["TLADesc"].Width = 500;
            //    ultraGridBOM.DisplayLayout.Bands[0].Columns["CompDesc"].Width = 800;
            //  //  ultraGridBOM.DisplayLayout.Bands[0].Columns["manufacturer"].Width = 200;
            //}
            Cursor.Current = Cursors.Default;
        }


    DataTable dtPopulate;
    private void bindData(string sql )
        {
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            try
            {
                SqlCommand cmd = new SqlCommand(sql, cn);
                cmd.CommandTimeout = 0;

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                cn.Close();
                ultraGridBOM.DataSource = ds.Tables[1];

                DataRow[] ErrorRows=ds.Tables[1].Select("Head_Message  like 'ERROR%'");
                if (ErrorRows.Count() > 0)
                    button1.Enabled = false;
                else
                    button1.Enabled = true;

                dtPopulate = ds.Tables[1];

             if (ds.Tables.Count>1)
                 if (ds.Tables[2].Rows.Count > 0)
                 {
                     lblAsmStage.Text = ds.Tables[2].Rows[0]["Head_AsmStage"].ToString();
                     this.lblFormFactor.Text = ds.Tables[2].Rows[0]["Head_FormFactor"].ToString();
                     this.lblPlanGroup.Text = ds.Tables[2].Rows[0]["Head_PlanGroup"].ToString();
                     this.lblPlatform.Text = ds.Tables[2].Rows[0]["Head_Platform"].ToString();
                     this.lblPartDesc.Text = ds.Tables[2].Rows[0]["Head_TLADesc"].ToString();
                 }
                lblNumRecKan.Text = "Num of Records :" + ds.Tables[1].Rows.Count.ToString();
            }
            catch (Exception exp)
            {
                MessageBox.Show("Error to populate spend tool. Please check the error" + exp.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            cn.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (radBacon.Checked)
            {
                string sqlinsert = "   insert into dbo.goo_spend_tool_bom(CM_part_number,google_part_number,parent_google_part_number,manf_part_number,manufacturer,description,quantity,unit_price,ext_price,moq,lead_time,neog_by_google,assm_by_google,tran_id)   SELECT CM_part_number,google_part_number,parent_google_part_number,manf_part_number,manufacturer,description,quantity,unit_price,ext_price,moq,lead_time,neog_by_google,assm_by_google,tran_id  FROM [MIMDIST].[dbo].[BaconBOM_for_SpendTool] ";
                string sql = "";

                //for (int i = 0; i < ultraGridBOM.Rows.Count; i++)
                //{
                //    string CM_part_number = ultraGridBOM.Rows[i].Cells["CM_part_number"].Value.ToString();
                //    string google_part_number = ultraGridBOM.Rows[i].Cells["google_part_number"].Value.ToString();
                //    string parent_google_part_number = ultraGridBOM.Rows[i].Cells["parent_google_part_number"].Value.ToString();
                //    string manf_part_number = ultraGridBOM.Rows[i].Cells["manf_part_number"].Value.ToString();
                //    string manufacturer = ultraGridBOM.Rows[i].Cells["manufacturer"].Value.ToString();
                //    string description = ultraGridBOM.Rows[i].Cells["description"].Value.ToString();
                //    string quantity = Convert.ToDecimal(ultraGridBOM.Rows[i].Cells["quantity"].Value).ToString();
                //    string unit_price = Convert.ToDecimal(ultraGridBOM.Rows[i].Cells["unit_price"].Value).ToString();
                //    string ext_price = Convert.ToDecimal(ultraGridBOM.Rows[i].Cells["ext_price"].Value).ToString();
                //    string moq = ultraGridBOM.Rows[i].Cells["moq"].Value.ToString();
                //    string lead_time = ultraGridBOM.Rows[i].Cells["lead_time"].Value.ToString();
                //    string neog_by_google = ultraGridBOM.Rows[i].Cells["neog_by_google"].Value.ToString();
                //    string assm_by_google = ultraGridBOM.Rows[i].Cells["assm_by_google"].Value.ToString();
                //    sql = sql + sqlinsert + " values ('" + CM_part_number + "','" + google_part_number + "','" + parent_google_part_number + "','" + manf_part_number + "','" + manufacturer + "','" + description + "','" + quantity + "','" + unit_price + "','" + ext_price + "','" + moq + "','" + lead_time + "','" + neog_by_google + "','" + assm_by_google + "')";
                //}

                sql = "  delete from dbo.[goo_spend_tool_bom] " + sqlinsert;
              
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlTransaction tran = cn.BeginTransaction();
                try
                {
                    SqlCommand cmd = new SqlCommand(sql, cn);
                    cmd.CommandTimeout = 0;
                    cmd.Transaction = tran;
                    cmd.ExecuteNonQuery();
                    tran.Commit();
                    MessageBox.Show("Spend tool table has been poplulated successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    runJob();
                    //Microsoft.SqlServer.Dts.Runtime.Application app = new Microsoft.SqlServer.Dts.Runtime.Application();
                    ////string pkgLocation;
                    
                    //// DTSExecResult pkgResults;

                    //string pkgLocation = @"\\mverp\f$\Integration\SupplierIntegration\SpendTool.dtsx";
                    //Package pkg = app.LoadPackage(pkgLocation, null);
                    //Microsoft.SqlServer.Dts.Runtime.DTSExecResult results = pkg.Execute();

                    //if (results == Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure)
                    //{
                    //    string err = "";
                    //    foreach (Microsoft.SqlServer.Dts.Runtime.DtsError local_DtsError in pkg.Errors)
                    //    {
                    //        string error = local_DtsError.Description.ToString();
                    //        err = err + error;
                    //    }
                    //}
                    //else if (results == Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success)
                    //{
                    //    MessageBox.Show("Package Executed Successfully....");
                    //}
                }
                catch (Exception exp)
                {
                    tran.Rollback();
                    MessageBox.Show("Not Saved. Please check the error" + exp.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                cn.Close();
            }
            else if (radShareDoc.Checked)
                {
                    create_file();
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

                    FilesResource.ListRequest lstreq = service.Files.List();
                    string Parentid = "";
                    Parentid = getGIDByName(Collection);
                string pageToken=null;

                if (Parentid == "")
                    MessageBox.Show("Parent not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                string fileid = "";

                do
                {


                    var result = getFilesByParentId(Parentid, pageToken);
                    foreach (var file in result.Files)
                    {

                        if (file.Name == TLA_PartNum)
                        {
                            fileid = file.Id;
                            break;

                        }
                    }
                    pageToken = result.NextPageToken;
                } while (pageToken != null);



                //    string fileid = getchildInFolder(service, Parentid);


                    try
                    {
                        Google.Apis.Drive.v3.Data.File body = new Google.Apis.Drive.v3.Data.File();
                        body.Name = TLA_PartNum;
                        body.Description = TLA_PartNum;
                        body.MimeType = "application/vnd.google-apps.spreadsheet";
                        body.Parents = new List<string> { Parentid }; 
                        //  string newTitle = System.IO.Path.GetFileName("Rack Delivery Plan" + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Year.ToString("00") + ".xlsx");
                    //    if (!String.IsNullOrEmpty(Parentid))
                      //  {
                      //      body.Parents = new List<ParentReference>() { new ParentReference() { Id = Parentid } };
                     //   }

                        //  filename = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\MIM_Menlo Demand TrackerTemplate.xlsx";// "E:\\Ruchi\\b2bdevMaintainGoogleDoc\\MaintainGoogleDoc\\MaintainGoogleDoc\\bin\\Debug\\MIM_Menlo Demand TrackerTemplate.xlsx";
                        filename = filename1.Replace("file:\\", " ");
                        // MessageBox.Show("filename=" + filename);
                        System.IO.Stream fileStream = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                        byte[] byteArray = new byte[fileStream.Length];
                        fileStream.Read(byteArray, 0, (int)fileStream.Length);
                        System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
                        if (fileid != "")
                        {
                            FilesResource.UpdateMediaUpload request1 = service.Files.Update(body, fileid, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
                            //request1.Convert = true;
                            request1.Upload();
                        }
                        else
                        {
                            FilesResource.CreateMediaUpload request2 = service.Files.Create(body, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
                         //   request2.Convert = true;
                            request2.Upload();
                        }
                        
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

        private void runJob()
        {
             SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlTransaction tran = cn.BeginTransaction();
                try
                {
                    SqlCommand cmd = new SqlCommand(" EXEC msdb.dbo.sp_start_job 'Internal-SpendTool' ", cn);
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
        }

        DriveService service;
        private void radShareDoc_CheckedChanged(object sender, EventArgs e)
        {
           
           
        }
        string fileid="";
     /*   public string getchildInFolder(DriveService service, String folderId)
        {

            fileid = "";
            ChildrenResource.ListRequest request = service.Children.List(folderId);

            do
            {
                try
                {
                    ChildList children = request.Execute();

                    foreach (ChildReference child in children.Items)
                    {
                        Google.Apis.Drive.v2.Data.File file1 = service.Files.Get(child.Id).Execute();
                        if (file1.Title == TLA_PartNum)
                        {
                            fileid = file1.Id; 
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

            return fileid;
        }
      * */
        Microsoft.Office.Interop.Excel.Worksheet sheet;
        Microsoft.Office.Interop.Excel.Sheets excelSheets;
        Microsoft.Office.Interop.Excel.Application excelApp1;
        Microsoft.Office.Interop.Excel.Workbook excelWorkbook;
        string filename1 = "";
        string Collection = "";
        string TLA_PartNum ="";
        string TLA_PartDesc ="";
      
        string filename = "";

      

        public void create_file()
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            try
            {

                String strQuery = "     select  Top 1 [SharedDocFolderName],TLA_PartNum,TLADesc  FROM [MIMDIST].[dbo].[BaconBOM_for_SharedDoc] where [SharedDocFolderName] is not null   SELECT  [AssemblyType] ,[AssyGPN],[CommodityCode] ,[SupplyType],[CompGPN] ,[LeadTime] ,[BOMQty] ,[UnitCost],[ExtCost],[CompDesc]    FROM [MIMDIST].[dbo].[BaconBOM_for_SharedDoc] order by sortorder,COMPGPN ";
                DataSet ds = getDataSet(strQuery);

                DataTable dt = dtPopulate.Clone();
                dt = ds.Tables[1];// ultraGridBOM.DataSource as DataTable;
                Collection =ds.Tables[0].Rows[0][0].ToString();
                TLA_PartNum = ds.Tables[0].Rows[0][1].ToString();
                TLA_PartDesc = ds.Tables[0].Rows[0][2].ToString();


                excelApp1 = new Microsoft.Office.Interop.Excel.Application();
                // string workbookPath = "E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\DemandTrackerSheet.xls";// Microsoft.Office.Interop.Excel.XlPlatform.xlWindows  
                filename = "\\\\mvfile\\MV REPORTS\\Master Data & Forms\\PlanningExecutionForms\\Templates\\shared Doc Costed BOM Template.xlsx";

                excelApp1 = new Microsoft.Office.Interop.Excel.Application();
                filename1 = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\" + TLA_PartNum + ".xlsx";//Menlo_Metrics_Kit_SnapshotTemplate.xlsx";
                          
                string workbookPath = filename.Replace("file:\\", "");
                excelWorkbook = excelApp1.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                excelWorkbook.SaveAs(filename1, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

                excelSheets = excelWorkbook.Worksheets;
             
                sheet = excelSheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                sheet.Cells[1, 2] = TLA_PartNum;
                sheet.Cells[2, 2] = TLA_PartDesc;
                updateSheet(dt);
                excelWorkbook.Save();
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

        public void updateSheet(DataTable dt)
        {

            int col = 0;
           
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    string data = dt.Rows[i][j].ToString();
                    sheet.Cells[i + 5, j + 1] = dt.Rows[i][j].ToString();
                }
            }
            excelWorkbook.Save();
        }

        private void ultraGridBOM_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
        {


            if (e.Cell.Column.ToString() == "CommodityCode")
            {
                if (e.Cell.Text == "SHEET METAL")
                    e.Cell.Row.Appearance.BackColor = Color.LightBlue;
                else
                    e.Cell.Row.Appearance.BackColor = Color.LightPink;
            }
        }

        private void ultraPanel1_PaintClient(object sender, PaintEventArgs e)
        {

        }

        private void ultraGridBOM_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
        {
          

	
	// Set row background color when EmployeeID = 1
    if (e.Row.Cells["Head_Message"].Value.ToString().StartsWith("ERROR"))
	{	
		e.Row.Appearance.BackColor = Color.Red;
	//	e.Row.Appearance.BackColor2 = Color.Red;
	//	e.Row.Appearance.BackGradientStyle = GradientStyle.Horizontal;
		e.Row.Appearance.ForeColor = Color.White;
	}
    else if (e.Row.Cells["Head_Message"].Value.ToString().StartsWith("WARNING"))
    {
        e.Row.Appearance.BackColor = Color.Red;
        //	e.Row.Appearance.BackColor2 = Color.Red;
        //	e.Row.Appearance.BackGradientStyle = GradientStyle.Horizontal;
        e.Row.Appearance.ForeColor = Color.White;
    }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            
                    Microsoft.SqlServer.Dts.Runtime.Application app = new Microsoft.SqlServer.Dts.Runtime.Application();
                    //string pkgLocation;
                    
                    // DTSExecResult pkgResults;

                string     pkgLocation =  @"\\mverp\f$\Integration\SupplierIntegration\SpendTool.dtsx";

                Package pkg=app.LoadPackage(pkgLocation, null);

                Microsoft.SqlServer.Dts.Runtime.DTSExecResult results = pkg.Execute();

                if (results == Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure)
                {
                    string err = "";
                    foreach (Microsoft.SqlServer.Dts.Runtime.DtsError local_DtsError in pkg.Errors)
                    {
                        string error = local_DtsError.Description.ToString();
                        err = err + error;
                    }
                }
                else if (results == Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success)
                {
                   MessageBox.Show("Package Executed Successfully....");
                }
        }
    }
}
