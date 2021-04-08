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

namespace Version3
{
    public partial class NewBuildPlan : Form
    {
      
        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        string constrTEST ="";// ConfigurationManager.ConnectionStrings["epicorTESTConnectionString"].ConnectionString;
  

       string googleshareDoc;
       public NewBuildPlan()
        {
            loaded = false;          
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
              request.Q = "modifiedTime > '2016-04-01T12:00:00' and (name contains 'MIM TLA Plan' or name contains  'MIM Build Plan')  ";
             //  request.Q = "name = 'MIM tracker'";
              //  string parent_id = "0B1auUUiGcbbBOEIyNjg0alkwc3c";
              //  request.Q = "'" + parent_id + "' in parents";
               request.Spaces = "Drive";
               request.Fields = "*";
               request.PageToken = pageToken;
               var result = request.Execute();

               foreach (var file in result.Files)
               {
                 //  if (file.CreatedTime>System.DateTime.Now.AddMonths(-1))
                 //  if (file.Name.StartsWith("MIM Build Plan ") || file.Name.StartsWith("MIM TLA Plan "))
                  // {
                   if (file.Name.IndexOf('[') > 0)
                   {
                       ListViewItem item = new ListViewItem(new string[2] { file.Name.Substring(0, file.Name.IndexOf('[') - 1), file.Id });
                       this.spreadsheetListView.Items.Add(item);
                   }
                   else
                   {
                       ListViewItem item = new ListViewItem(new string[2] { file.Name, file.Id });
                       this.spreadsheetListView.Items.Add(item);
                   }

                  // }
               }
               pageToken = result.NextPageToken;
           } while (pageToken != null);


           do
           {
               var request = service.Files.List();
               string title = "GIG-MIM build plan";
               request.Q = "modifiedTime > '2018-04-01T12:00:00' and (name contains 'NPI TLA Plan' or name contains  'NPI Build Plan')  ";
               //  request.Q = "name = 'MIM tracker'";
               //  string parent_id = "0B1auUUiGcbbBOEIyNjg0alkwc3c";
               //  request.Q = "'" + parent_id + "' in parents";
               request.Spaces = "Drive";
               request.Fields = "*";
               request.PageToken = pageToken;
               var result = request.Execute();

               foreach (var file in result.Files)
               {
                   //  if (file.CreatedTime>System.DateTime.Now.AddMonths(-1))
                   //  if (file.Name.StartsWith("MIM Build Plan ") || file.Name.StartsWith("MIM TLA Plan "))
                   // {
                   if (file.Name.IndexOf('[') > 0)
                   {
                       ListViewItem item = new ListViewItem(new string[2] { file.Name.Substring(0, file.Name.IndexOf('[') - 1), file.Id });
                       this.spreadsheetListView.Items.Add(item);
                   }
                   else
                   {
                       ListViewItem item = new ListViewItem(new string[2] { file.Name, file.Id });
                       this.spreadsheetListView.Items.Add(item);
                   }

                   // }
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
            
           
        string filename;
        DriveService service;

   /*     //get spreadsheet buildplan from google shared doc
        void setSpreadsheetListView()
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
            
            string fileid = "";
            string Parentid = "";
            FilesResource.ListRequest lstreq = service.Files.List();
         
            do
            {
                try
                {

                    lstreq.Q = request.Q = "fullText contains '" + title + "'"; "title contains 'MIM TLA Plan '  or title contains 'MIM Build Plan '";
                  //  lstreq.Fields = "title/id";
                      FileList fileList = lstreq.Execute();                   

                    foreach (Google.Apis.Drive.v3.Data.File fileitem in fileList)
                    {
                        if (fileitem.Title.StartsWith("MIM Build Plan ") || fileitem.Title.StartsWith("MIM TLA Plan "))//"MIM NPI:" + projectName)
                        {
                            try
                            {
                                ListViewItem item = new ListViewItem(new string[2] { fileitem.Title.Substring(0, fileitem.Title.IndexOf('[') - 1), fileitem.Id });
                                this.spreadsheetListView.Items.Add(item);
                            }
                            catch { }
                       
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


        //   string file_id = getchildInFolder(service, Parentid);           
         
        }
    * */
        //get all worksheet for the selected buildplan put into worksheet Listview
  
        DataTable table = new DataTable();

   /*     public  string  getchildInFolder(DriveService service,   String folderId)
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
                        if (file1.Title.StartsWith("Build Plan"))
                        {
                            ListViewItem item = new ListViewItem(new string[2] { file1.Title, child.Id });
                          //  googleshareDoc = file1.Title;
                            this.spreadsheetListView.Items.Add(item);
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
*/
        string sheetName = "";
        string dbCol = "";
 
        Form frm;
        private void showMessage()
        {
            frm=new Form();
            frm.StartPosition = FormStartPosition.CenterScreen;
            Label lbl=new Label();
            lbl.Width = 200;
            frm.Size = new Size(200, 100);
            lbl.Text="Please wait.. Grid is loading";
            frm.Controls.Add(lbl);
            frm.ControlBox = false;
            frm.ShowInTaskbar = false;
            frm.ShowDialog(); 
                  
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


        Thread t;
       
        private void button1_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (spreadsheetListView.SelectedItems.Count > 0)
            {

                filename = spreadsheetListView.SelectedItems[0].SubItems[1].Text;
                googleshareDoc = spreadsheetListView.SelectedItems[0].SubItems[0].Text;
                //  fileId = spreadsheetListView.SelectedItems[0]..Text;
              //  Google.Apis.Drive.v3.Data.File file1 = service.Files.Get(filename).Execute();
                
                saveTo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\currentBuildPlan.xlsx";
                saveTo = saveTo.Replace("file:\\", "");
                if (this.downloadFile(service, filename, saveTo))
                    showInGrid();
            }
            Cursor.Current = Cursors.Default;
        }
        
        static string ExtractNumbers(string expr)
        {
            return string.Join(null, System.Text.RegularExpressions.Regex.Split(expr, "[^\\d]"));
        }
        
       
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

        string saveTo = "";
        private void spreadsheetListView_SelectedIndexChanged(object sender, EventArgs e)
        {
          
        }
        
        public void showInGrid()
        {
            System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source="+saveTo+";Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1\";");
            MyConnection.Open();
            DataTable dtExcelSchema;
            dtExcelSchema = MyConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
     //       MessageBox.Show(dtExcelSchema.Rows.Count.ToString());
            DataRow[] drSheet = null;
            drSheet = dtExcelSchema.Select("TABLE_NAME like '%MIM Build Plan%'");
             if (drSheet.Length>0)
            {
                string sheetName = drSheet[0]["TABLE_NAME"].ToString();
                // sheetName = "MIM Build Plan 2016-07-13 [[Enormous Sabre-Toothed Tiger]]";

                //      MessageBox.Show("Sheet Name is " + dtExcelSchema.Rows[0]["TABLE_NAME"].ToString() + "  " + dtExcelSchema.Rows[1]["TABLE_NAME"].ToString() + "   " + dtExcelSchema.Rows[2]["TABLE_NAME"].ToString());
                System.Data.OleDb.OleDbDataAdapter MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [" + sheetName + "]", MyConnection);
                MyCommand.TableMappings.Add("Table", "TestTable");
                System.Data.DataSet DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                DataTable myTable = DtSet.Tables[0];
                DataTable newtab = new DataTable();
                for (int colind = 0; colind < myTable.Columns.Count; colind++)
                    if (myTable.Rows[0][colind].ToString().Trim() != "")
                 {
                    newtab.Columns.Add(new DataColumn(myTable.Rows[0][colind].ToString()));
                 }
                myTable.Rows[0].Delete();
                myTable = myTable.Select("F1 <>''").CopyToDataTable();

                foreach (DataRow dr in myTable.Rows)
                {
                    DataRow drnew = newtab.NewRow();

                    for (int j = 0; j < newtab.Columns.Count; j++)
                        drnew[j] = dr[j];
                    newtab.Rows.Add(drnew);

                }
                dataGridView1.DataSource = newtab;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F1"].Width = 80;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F2"].Width = 70;+
                //dataGridView1.DisplayLayout.Bands[0].Columns["F3"].Width = 100;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F4"].Width = 140;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F5"].Width = 30;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F6"].Width = 80;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F7"].Width = 80;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F8"].Width = 350;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F9"].Width = 50;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F10"].Width = 80;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F11"].Width = 40;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F12"].Width = 100;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F13"].Width = 90;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F14"].Width = 50;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F15"].Width = 50;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F16"].Width = 50;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F17"].Width = 50;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F18"].Width = 50;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F19"].Width = 120;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F20"].Width = 120;
                //dataGridView1.DisplayLayout.Bands[0].Columns["F21"].Width = 80;
                //MyCommand.Fill(DtSet);

                //dataGridView1.DataSource = DtSet.Tables[0];
                lblNumOfRow.Text = "Numbers of rows : " + newtab.Rows.Count.ToString();
            }
            MyConnection.Close();      

        }

    /*    public Boolean downloadFile(DriveService _service, File _fileResource, string _saveTo)
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        int selectedRowIndex = 0;
        object[] Datarow1;   
        int strlen = 0;

        //upload build plan
       private void btnUpload_Click(object sender, EventArgs e)
       {
            Cursor.Current = Cursors.Hand;
            getHeaderMapping("MIM Build PLan");
            getHeaderMapping("MIM TLA PLan");
           // getHeaderMapping("MIM NPI PLan");
            MessageBox.Show("BuildPlan has been  processed successfully");    
            Cursor.Current = Cursors.Default;       
      }

       public bool UploadedBuildPlan(string planname)
       {        
           try
           {
               string sqlPlanExist = " select * from m_scheduler_filedata where filename='" + planname + "'";
               SqlConnection cnPlanExist = new SqlConnection(constr);
               cnPlanExist.Open();
               SqlCommand cmdPlanExist = new SqlCommand(sqlPlanExist, cnPlanExist);
               SqlDataAdapter daPlanExist = new SqlDataAdapter(cmdPlanExist);
               DataSet  dsPlanExist = new DataSet();
               daPlanExist.Fill(dsPlanExist);
               if (dsPlanExist.Tables.Count > 0)
                   if (dsPlanExist.Tables[0].Rows.Count > 0)
                       return true;
           }
           catch (Exception e)
           {
               MessageBox.Show("Error is " + e.Message.ToString());
           }                

           return false;
       }

       DataSet dsMappingTable;
       DataTable table1;

       private void getHeaderMapping(string planname)
       {
           string plandatename="";
           plandatename=googleshareDoc;
          /* if (googleshareDoc.ToUpper().Contains("TLA PLAN") &&  ! UploadedBuildPlan(googleshareDoc.ToUpper().Replace("TLA", "BUILD")))
           {               
                MessageBox.Show("BUILD PLAN hasn't been processed. Please first process BUILD PLAN then TLA PLAN", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
           }
           else{*/
              string column_names = "";
              string sqlstrHeaderMapping = "SELECT TableName, FieldName,SharedDocLike,DocHeaderName,DataType FROM [BuildPlanP_FieldMapping] where SharedDocLike='" + planname + "'";//+ googleshareDoc.Substring(0, googleshareDoc.IndexOf("Plan") + 4) + "'";
               SqlConnection cn = new SqlConnection(constr);
               cn.Open();
               SqlCommand cmd = new SqlCommand(sqlstrHeaderMapping, cn);
               SqlDataAdapter da = new SqlDataAdapter(cmd);
               dsMappingTable = new DataSet();
               da.Fill(dsMappingTable);
               table1 = dataGridView1.DataSource as DataTable;
               if (planname == "MIM TLA PLan")
               {
                   
                   DataView view = new DataView(table1);
                   table = view.ToTable(true, "BuildID", "tlaPartNumber", "tlaDescription", "tlarsqty", "cluster", "subcontinent", "devBOM", "metro", "npiStage", "tlatotalQty", "orderDate", "orderType", "projectCode", "tlaremainingQty");
               }
               else
               {
                   table = table1;
               }

               string tablename = dsMappingTable.Tables[0].Rows[0]["TableName"].ToString();
               for (int i = 0; i < dsMappingTable.Tables[0].Rows.Count; i++)
               {
                   column_names = column_names + dsMappingTable.Tables[0].Rows[i]["FieldName"].ToString() + ",";
               }
               column_names = column_names.Substring(0, column_names.Length - 1);
               string sqlSaveFixed = "delete from " + tablename;
               for (int igd = 0; igd < table.Rows.Count; igd++)
               {
                   // if (table.Rows[igd]["BuildID"].ToString().Trim() != "build_id")
                   //{
                   string val = "";
                   int qty = 0;
                   string colname ="";
                   sqlSaveFixed = sqlSaveFixed + "  insert into " + tablename + "( " + column_names + ")  values('";

                   for (int imap = 0; imap < dsMappingTable.Tables[0].Rows.Count; imap++)
                       if (dsMappingTable.Tables[0].Rows[imap]["DataType"].ToString().ToUpper() == "INT")
                       {
                           try
                           {
                               colname = dsMappingTable.Tables[0].Rows[imap]["DocHeaderName"].ToString();
                                val = table.Rows[igd][dsMappingTable.Tables[0].Rows[imap]["DocHeaderName"].ToString()].ToString();
                               qty = Convert.ToInt32(val.Replace(",", ""));
                           }
                           catch (Exception e)
                           {
                              
                               qty = 0;
                           }
                           sqlSaveFixed = sqlSaveFixed + qty + "','";
                       }
                       else
                           sqlSaveFixed = sqlSaveFixed + table.Rows[igd][dsMappingTable.Tables[0].Rows[imap]["DocHeaderName"].ToString()].ToString().Replace("'", "").ToUpper() + "','";
                   sqlSaveFixed = sqlSaveFixed.Substring(0, sqlSaveFixed.Length - 2) + ")";//4609_23857
                   // }
               }

               SqlConnection cnSave = new SqlConnection(constr);
               cnSave.Open();
               SqlTransaction tran = cn.BeginTransaction();
               try
               {
                   SqlCommand cmdSave = new SqlCommand(sqlSaveFixed, cnSave);
                   cmdSave.ExecuteNonQuery();
                   tran.Commit();
               }
               catch (Exception e)
               {
                   tran.Rollback();
                   MessageBox.Show("Error is " + e.Message.ToString());
               }
               // SqlDataAdapter da = new SqlDataAdapter(cmd);
               //  dsMappingTable = new DataSet();
               //  da.Fill(dsMappingTable);
               getFileID(tablename);
              
      // }
           
       }

       private void  getFileID(string tablename)
       {
          string num = ExtractNumbers(spreadsheetListView.SelectedItems[0].SubItems[0].Text);
          DateTime filedate = DateTime.Now.Date;// Convert.ToDateTime(num.Substring(0, 4) + "-" + num.Substring(4, 2) + "-" + num.Substring(6, 2));
           try
           {
               string sqlFileID ="	update " + tablename + " set filename='" + spreadsheetListView.SelectedItems[0].SubItems[0].Text + "'"  ;
               if (tablename.Contains("TLA"))

                     sqlFileID = sqlFileID + "  exec [dbo].[BuildPlanP_TLA_ImportToStage_SP]"  ;
               else
                    sqlFileID = sqlFileID + "  exec [dbo].[BuildPlanP_DEL_ImportToLive_SP]"  ;   
             
               SqlConnection cnFileID = new SqlConnection(constr);
               cnFileID.Open();
               SqlCommand cmdFileID = new SqlCommand(sqlFileID, cnFileID);
               cmdFileID.ExecuteNonQuery();
           }
           catch (Exception e)
           {
               MessageBox.Show("Error is " + e.Message.ToString());
           }                  
       }

       //public void getExcelcolumnNameNSP()
       //{
       //    SqlConnection cn = new SqlConnection(constr);
       //    cn.Open();
       //    SqlCommand cmd = new SqlCommand("select * from  m_nsp_header_mapping", cn);
       //    SqlDataAdapter da = new SqlDataAdapter(cmd);
       //    dsMappingTable = new DataSet();
       //    da.Fill(dsMappingTable);
       //    string[] colDatatype = new string[100];
       //    char[] colConstant = new char[100];
       //    string notMatchColumn = "";
       //    string[] colName = new string[100];

       //    string[] tblcolName = new string[100];
       //    int j = 0;
       //    System.Data.DataTable dt = new System.Data.DataTable();
       //    dt.Clear();
       //    dbCol = "";
       //   string dbFixedcol = "";
       //    for (int i = 2; i < dsMappingTable.Tables[0].Rows.Count; i++)
       //    {
       //        if (i != 2)
       //            dbCol = dbCol + ",";


       //        dbCol = dbCol + dsMappingTable.Tables[0].Rows[i][1].ToString();

       //        //if (dsMappingTable.Tables[0].Rows[i][5].ToString().ToUpper() == "FIXED")
       //        //{
       //        //    if (i != 2)
       //        //        dbFixedcol = dbFixedcol + ",";
       //        //    dbFixedcol = dbFixedcol + dsMappingTable.Tables[0].Rows[i][1].ToString();
       //        //}

       //        dt.Columns.Add(dsMappingTable.Tables[0].Rows[i][2].ToString().ToUpper());
       //        tblcolName[j] = dsMappingTable.Tables[0].Rows[i][1].ToString().ToLower();
       //        colName[j] = dsMappingTable.Tables[0].Rows[i][2].ToString().ToLower();
       //        colDatatype[j] = dsMappingTable.Tables[0].Rows[i][3].ToString().ToLower();
       //        j++;
       //    }

       //    //dt.Columns.Add("delete");
       //    // MessageBox.Show(dt.Columns.Count.ToString());
       //    foreach (DataRow sourcerow in table.Rows)
       //    {
       //        DataRow destRow = dt.NewRow();
       //        for (int k = 0; k < j; k++)
       //        {

       //            try
       //            {
       //                if (colDatatype[k] == "integer" && (sourcerow[colName[k]].ToString() == "" || sourcerow[colName[k]] == null))
       //                    sourcerow[colName[k]] = 0;
       //                else if (colDatatype[k] == "varchar" && (sourcerow[colName[k]].ToString().Trim() == "" || sourcerow[colName[k]] == null || sourcerow[colName[k]].ToString().Trim() == "0"))
       //                    sourcerow[colName[k]] = "";
       //                destRow[colName[k]] = sourcerow[colName[k]];
       //            }
       //            catch (Exception ex)
       //            {
       //                MessageBox.Show(colName[k] + "  does not match");

       //                //if (!notMatchColumn.Contains(colName[k]))
       //                //{
       //                //    if (notMatchColumn != "")
       //                //        notMatchColumn = notMatchColumn + ",";
       //                //    notMatchColumn = notMatchColumn + colName[k];
       //                //}
       //            }
       //                           }

       //        dt.Rows.Add(destRow);
       //    }
       //}

        //Bulk insert the DataTable data to sql server table
       private void WriteToDatabase()
       {
           // get your connection string
           string connString = "";
           // connect to SQL
           using (SqlConnection connection =
                   new SqlConnection(constr))
           {
               // make sure to enable triggers
               // more on triggers in next post
               SqlBulkCopy bulkCopy =
                   new SqlBulkCopy
                   (
                   connection,
                   SqlBulkCopyOptions.TableLock |
                   SqlBulkCopyOptions.FireTriggers |
                   SqlBulkCopyOptions.UseInternalTransaction,
                   null
                   );

               // set the destination table name
               bulkCopy.DestinationTableName = "m_temp_build_plan";
               connection.Open();

               // write the data in the "dataTable"
               bulkCopy.WriteToServer(this.table);
               connection.Close();
           }
           // reset        
       }

       private void lstWorksheet_SelectedIndexChanged(object sender, EventArgs e)
       {
           
       }
     

       private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
       {

       }

     
    }

    //public class ListItem1
    //{
    //    private string id = string.Empty;
    //    private string name = string.Empty;
    //    public ListItem1(string sid, string sname)
    //    {
    //        id = sid;
    //        name = sname;
    //    }
        
    //    public override string ToString()
    //    {
    //        return this.name;
    //    }

    //    public string ID
    //    {
    //        get
    //        {
    //            return this.id;
    //        }
    //        set
    //        {
    //            this.id = value;
    //        }
    //    }

    //    public string Name
    //    {
    //        get
    //        {
    //            return this.name;
    //        }
    //        set
    //        {
    //            this.name = value;
    //        }
    //    }
       
    //}

    
}
