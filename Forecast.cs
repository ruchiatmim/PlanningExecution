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



namespace Version3

{
    public partial class Forecast : Form
    {
      
        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        string constrTest = ConfigurationManager.ConnectionStrings["epicorTESTConnectionString"].ConnectionString;
        string constrMVAPPS = ConfigurationManager.ConnectionStrings["MVAPPSDataAnalysisConnectionString"].ConnectionString;
        string constrMVReports = ConfigurationManager.ConnectionStrings["MVReportsShopfloorConnectionString"].ConnectionString;
        //  string constr = "Data Source=Epicor;Initial Catalog=MIMDIST;user id=sa;password=mimi~100;Pooling=true;Max Pool Size=100;Min Pool Size=1;";

        public Forecast()
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
                string title = "Current NSP";
                // request.Q = "fullText contains '" + title + "'";
                string parent_id = "0Bw-Ha0rHFFlEUTRCUUJDQnhZUnc";
                request.Q = "'" + parent_id + "' in parents";
                request.Spaces = "Drive";
                request.Fields = "nextPageToken, files(id, name)";
                request.PageToken = pageToken;

                var result = request.Execute();
                foreach (var file in result.Files)
                {
                    if (file.Name.StartsWith("Net Supply Plan"))
                    {
                        ListViewItem item = new ListViewItem(new string[2] { file.Name, file.Id });
                        googleshareDoc = file.Name;
                        this.spreadsheetListView1.Items.Add(item);

                    }
                }
                pageToken = result.NextPageToken;
            } while (pageToken != null);

            do
            {
                var request = service.Files.List();
                string title = "Weekly NSP";
               // request.Q = "fullText contains '" + title + "'";
                 string parent_id = "0B39rSdVkAq1ZRkZxUWlnZDJWWkU";
                 request.Q = "'" + parent_id + "' in parents";
                request.Spaces = "Drive";
                request.Fields = "nextPageToken, files(id, name)";
                request.PageToken = pageToken;

                var result = request.Execute();
                foreach (var file in result.Files)
                {
                    if (file.Name.StartsWith("Net Supply Plan"))
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

        DriveService service;
        string fileid = "";
        string Parentid = "";
        string Collection = "MIM Menlo Demand Tracker";
        string docName = "Net Supply Plan";
        string googleshareDoc = "";
        string folder_id = "";
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
            MessageBox.Show("start3333");
            do
            {
                try
                {
                    lstreq.Q = " title='Weekly NSP'";
                    FileList fileList = lstreq.Execute();
                    foreach (Google.Apis.Drive.v2.Data.File fileitem in fileList.Items)
                    {
                       
                           // ListViewItem item = new ListViewItem(new string[2] { fileitem.Title, fileitem.Id });

                            folder_id = fileitem.Id;
                            goto findChild;
                           // googleshareDoc = fileitem.Title;
                           // this.spreadsheetListView1.Items.Add(item);
                      // }
                           
                    }
                    lstreq.PageToken = fileList.NextPageToken;
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                    lstreq.PageToken = null;
                }
            } while (!String.IsNullOrEmpty(lstreq.PageToken));

        findChild:
            getALLchildInFolder(service, folder_id);
        }

        string saveTo = "";

     //   string googleshareDoc = "";
        public void getALLchildInFolder(DriveService service, String folderId)
        {
            this.spreadsheetListView1.Items.Clear();
            ChildrenResource.ListRequest request = service.Children.List(folderId);
            do
            {
                try
                {
                    request.Q = " title contains 'Net Supply Plan'";
                    ChildList children = request.Execute();

                    foreach (ChildReference child in children.Items)
                    {
                        Google.Apis.Drive.v2.Data.File file1 = service.Files.Get(child.Id).Execute();

                       // if (file1.Title.Contains("Net Supply Plan"))
                        //{
                            ListViewItem item = new ListViewItem(new string[2] { file1.Title, child.Id });
                            googleshareDoc = file1.Title;
                            this.spreadsheetListView1.Items.Add(item);
                        //}

                    }
                    request.PageToken = children.NextPageToken;
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                    request.PageToken = null;
                }
            } while (!String.IsNullOrEmpty(request.PageToken));


        }    
  


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
         * */

        string table_name = "";
        string ALLColumns = "";
        string saveTo = "";
        public void showInGrid(string table_name)
        {
            System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + saveTo + "';Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1\";");
            System.Data.OleDb.OleDbDataAdapter MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [" + table_name + "$]", MyConnection);
            MyCommand.TableMappings.Add("Table", "TestTable");
            System.Data.DataSet DtSet = new System.Data.DataSet();
            MyCommand.Fill(DtSet);
            DataTable dtShareColl = DtSet.Tables[0];
            int col=0;
            table = new DataTable();
           for ( col = 0;col < dtShareColl.Columns.Count;col++)
           {
               table.Columns.Add(dtShareColl.Rows[0][col].ToString());
           }
           dtShareColl.Rows.RemoveAt(0);
           System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("[0-9]");
           for (int row = 0; row < dtShareColl.Rows.Count; row++)
           {
                DataRow dr=table.NewRow();
                for (col = 0; col < dtShareColl.Columns.Count; col++)
                {

                  
                    try
                    {
                        if (table.Columns[col].ColumnName.ToUpper().Replace("_", " ") == "PLANNING ITEM" && dtShareColl.Rows[row][col].ToString().StartsWith("7"))
                        {
                            if (regex.Match(dtShareColl.Rows[row][col].ToString().Substring(0, 1).ToString()).Success)
                                dr[col] = "0" + dtShareColl.Rows[row][col].ToString();
                            else
                                dr[col] = dtShareColl.Rows[row][col].ToString();
                        }
                        else
                            dr[col] = dtShareColl.Rows[row][col].ToString().Replace("RACKSET", "RS"); ;

                    }
                    catch
                    {
                        dr[col] = dtShareColl.Rows[row][col].ToString();
                    }                    
                }
                   
                table.Rows.Add(dr);
           }
         //   table = DtSet.Tables[0];
            MyConnection.Close();

            for (int i = 0; i < dgvNewforecast.ColumnCount; i++)
                   dgvNewforecast.Columns[i].Width = 200;

         //   dgvNewforecast.DataSource = table;
                string tablename = "m_nsp_plan";
                if (radioButton2.Checked)
                    tablename = "m_nsp_plan";
                if (radioButton3.Checked)
                    tablename = "m_nsp_deployed_raw";
          
                getExcelcolumnNameNSP(tablename);              
                lblNumOfRec.Text = "Number of Records :" + dtShareColl.Rows.Count.ToString();

            }
     

        public void getExcelcolumnNameNSP(string table_name)
        {
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();

            string qry="  select * from  m_nsp_header_mapping where table_name ='" + table_name + "'   order by ord  ";
            SqlCommand cmd = new SqlCommand(qry, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            dsMappingTable = new DataSet();
            da.Fill(dsMappingTable);
            string[] colDatatype = new string[100];
            char[] colConstant = new char[100];
            string notMatchColumn = "";
            string[] colName = new string[100];

            string[] tblcolName = new string[100];
            int j = 0;
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Clear();
            dbCol = "";
            dbSharedCol = "";
            dbFixedcol = "";
            dbDateType = "";
            for (int i = 2; i < dsMappingTable.Tables[0].Rows.Count; i++)
            {
                if (i != 2)
                {
                    dbCol = dbCol + ",";
                    dbSharedCol = dbSharedCol + ",";
                    dbDateType = dbDateType + ",";
                }

                dbCol = dbCol + dsMappingTable.Tables[0].Rows[i][1].ToString();
                dbSharedCol = dbSharedCol + dsMappingTable.Tables[0].Rows[i][2].ToString();
                dbDateType=dbDateType +  dsMappingTable.Tables[0].Rows[i][3].ToString().ToLower();

                //if (dsMappingTable.Tables[0].Rows[i][5].ToString().ToUpper() == "FIXED")
                //{
                //    if (i != 2)
                //        dbFixedcol = dbFixedcol + ",";
                //    dbFixedcol = dbFixedcol + dsMappingTable.Tables[0].Rows[i][1].ToString();
                //}
                
                dt.Columns.Add(dsMappingTable.Tables[0].Rows[i][1].ToString().ToUpper());
                tblcolName[j] = dsMappingTable.Tables[0].Rows[i][1].ToString().ToLower();
                colName[j] = dsMappingTable.Tables[0].Rows[i][2].ToString().ToLower();
                colDatatype[j] = dsMappingTable.Tables[0].Rows[i][3].ToString().ToLower();
                j++;
            }

            //dt.Columns.Add("delete");
            // MessageBox.Show(dt.Columns.Count.ToString());
            foreach (DataRow sourcerow in table.Rows)
            {
                DataRow destRow = dt.NewRow();
                for (int k = 0; k < j; k++)
                {

                    try
                    {
                        if (colDatatype[k] == "integer" && (sourcerow[colName[k]].ToString() == "" || sourcerow[colName[k]] == null))
                            sourcerow[colName[k]] = 0;
                        else if (colDatatype[k] == "varchar" && (sourcerow[colName[k]].ToString().Trim() == "" || sourcerow[colName[k]] == null || sourcerow[colName[k]].ToString().Trim() == "0"))
                            sourcerow[colName[k]] = "";
                        destRow[tblcolName[k]] = sourcerow[colName[k]];
                    }
                    catch (Exception ex)
                    {
                      
                    }
                }

                dt.Rows.Add(destRow);
            }
            dgvNewforecast.DataSource = dt;
          
        }


    /*    public static void PrintFilesInFolder(DriveService service,        String folderId)
        {
            ChildrenResource.ListRequest request = service.Children.List(folderId);
            do
            {
                try
                {
                    ChildList children = request.Execute();
                    foreach (ChildReference child in children.Items)
                    {
                        Console.WriteLine("File Id: " + child.Id);
                    }
                    request.PageToken = children.NextPageToken;
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                    request.PageToken = null;
                }
            } while (!String.IsNullOrEmpty(request.PageToken));
        }
        
        public static bool isFileInFolder(DriveService service, String folderId, String fileId)
        {
            try
            {
                service.Children.Get(folderId, fileId).Execute();
            }

            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
              //  throw;
                return false;
            }
            return true;
        }*/
      
        DataTable table = new DataTable();

        string sheetName = "";     
        string dbCol = "";
        string dbSharedCol = "";
        string dbFixedcol = "";
        string dbDateType = "";
        DataSet dsMappingTable;


        //public void getExcelcolumnNameNSP(string table_name)
        //{
        //    SqlConnection cn = new SqlConnection(constr);
        //    cn.Open();

        //    string qry="  select * from  m_nsp_header_mapping where table_name ='" + table_name + "'   order by ord  ";
        //    SqlCommand cmd = new SqlCommand(qry, cn);
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
        //    dbSharedCol = "";
        //    dbFixedcol = "";
        //    dbDateType = "";
        //    for (int i = 2; i < dsMappingTable.Tables[0].Rows.Count; i++)
        //    {
        //        if (i != 2)
        //        {
        //            dbCol = dbCol + ",";
        //            dbSharedCol = dbSharedCol + ",";
        //            dbDateType = dbDateType + ",";
        //        }

        //        dbCol = dbCol + dsMappingTable.Tables[0].Rows[i][1].ToString();
        //        dbSharedCol = dbSharedCol + dsMappingTable.Tables[0].Rows[i][2].ToString();
        //        dbDateType=dbDateType +  dsMappingTable.Tables[0].Rows[i][3].ToString().ToLower();

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
        //              //   MessageBox.Show(colName[k] + "  does not match");


        //                //if (!notMatchColumn.Contains(colName[k]))
        //                //{
        //                //    if (notMatchColumn != "")
        //                //        notMatchColumn = notMatchColumn + ",";
        //                //    notMatchColumn = notMatchColumn + colName[k];
        //                //}
        //            }

        //        }

        //        dt.Rows.Add(destRow);
        //    }
        //    //**************************************************************update**********************************************/
        //    // DataView dvData = new DataView(dt);
        //    // dvData.RowFilter = "actiontype= 'BUILD' or status='FLEXIBILITY BUFFER'";
        //    // dt = dvData.ToTable();
        //    //**************************************************************update**********************************************/
        //}

        //public void getExcelcolumnName()
        //{
        //    SqlConnection cn = new SqlConnection(constr);
        //    cn.Open();

        //    SqlCommand cmd = new SqlCommand("select * from  m_serv_forecast_header_mapping", cn);
        //    SqlDataAdapter da = new SqlDataAdapter(cmd);
        //     dsMappingTable = new DataSet();
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
        //    dbFixedcol = "";
        //    for (int i = 2; i < dsMappingTable.Tables[0].Rows.Count; i++)
        //    {
        //        if (i != 2)
        //            dbCol = dbCol + ",";
                    
                
        //            dbCol = dbCol + dsMappingTable.Tables[0].Rows[i][1].ToString();

        //            if (dsMappingTable.Tables[0].Rows[i][5].ToString().ToUpper() == "FIXED")
        //            {
        //                if (i!=2)
        //                dbFixedcol = dbFixedcol + ",";
        //                dbFixedcol = dbFixedcol + dsMappingTable.Tables[0].Rows[i][1].ToString();
        //            }              
        //        dt.Columns.Add(dsMappingTable.Tables[0].Rows[i][2].ToString().ToUpper());
        //        tblcolName[j] = dsMappingTable.Tables[0].Rows[i][1].ToString().ToLower();
        //        colName[j] = dsMappingTable.Tables[0].Rows[i][2].ToString().ToLower();
        //        colDatatype[j] = dsMappingTable.Tables[0].Rows[i][4].ToString().ToLower();
              
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
        //                // MessageBox.Show(colName[k] + "  does not match");


        //                if (!notMatchColumn.Contains(colName[k]))
        //                {
        //                    if (notMatchColumn != "")
        //                        notMatchColumn = notMatchColumn + ",";
        //                    notMatchColumn = notMatchColumn + colName[k];
        //                }
        //            }

        //        }

        //        dt.Rows.Add(destRow);
        //    }
        //  //**************************************************************update**********************************************/
        //  // DataView dvData = new DataView(dt);
        //  // dvData.RowFilter = "actiontype= 'BUILD' or status='FLEXIBILITY BUFFER'";
        //  // dt = dvData.ToTable();
        //  //**************************************************************update**********************************************/

        //    //  MessageBox.Show(dt.Rows.Count.ToString());  
        //    if (notMatchColumn == "nic_type,nic_qty" || notMatchColumn == "")
        //    {

        //        try
        //        {
        //            pnl.Visible = false;
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message.ToString());
        //        }
        //        panel1.Controls.Remove(pnl);
        //        dataGridView1.DataSource = null;
        //        dataGridView1.DataSource = dt;
        //        lblRowCount.Text = "Rows are :" + dt.Rows.Count.ToString();
        //        dataGridView1.Columns[dataGridView1.ColumnCount - 1].Width = 200;
        //    }
        //    else
        //    {
        //        dt.Clear();
        //        dataGridView1.DataSource = null;
        //        showNonMatchcolList(notMatchColumn);
        //    }

        //}
        //GroupBox pnl = new GroupBox();
        //private void showNonMatchcolList(string notMatchColumn)
        //{
        //    String[] arrnot = notMatchColumn.Split(',');
        //    int cnt = arrnot.Count();

        //    pnl.ForeColor = Color.White;
        //    pnl.Text = "NON MATCHING COLUMNS";
        //    pnl.Width = 600;
        //    pnl.Location = new Point(1250, 3);
        //    int ypt = 20;
        //    int xpt = 10;
        //    for (int i = 0; i < cnt; i++)
        //    {
        //        Label lbl = new Label();
        //        lbl.Text = arrnot[i].ToUpper();
        //        lbl.Width = 200;
        //        if (i != 0)
        //        {
        //            ypt = ypt + 25;
        //        }
        //        if ((i % 3 == 0) && (i != 0))
        //        {
        //            ypt = 20;
        //            xpt = xpt + 200;
        //        }

        //        lbl.Name = "lbl" + i;
        //        lbl.Location = new Point(xpt, ypt);
        //        pnl.Controls.Add(lbl);
        //    }
        //    //  MessageBox.Show(pnl.Controls.Count.ToString());
        //    panel1.Controls.Add(pnl);
        //    pnl.Visible = true;

        //}

        //       private void updateExempt()
        //       {
        //           String[] arrnot = notMatchColumn.Split(',');
        //           int cnt = arrnot.Count();
        //           DataTable dt = dataGridView1.DataSource as DataTable;
        //           string sqlSave = " update m_serv_forecast_header_mapping set exempt='Y'
        //";
        //           SqlConnection cn = new SqlConnection(constr);
        //           cn.Open();

        //           SqlCommand cmd = new SqlCommand(sqlSave, cn);
        //           cmd.ExecuteNonQuery();
        //           cn.Close();
        //       }


        //private void btnUpload_Click(object sender, EventArgs e)
        //{
           
        //    if (WorkSheet.Contains("Firm Demand"))
        //    {
        //        DataTable dtBuild = dataGridView1.DataSource as DataTable;
        //        string sqlDelete = "delete from m_serv_forecast where status<>'FLEXIBILITY BUFFER' ";
        //        string sqlSave = "";
        //        string sqlSaveFixed = "";
        //        /*    DataView dvData = new DataView(dt);
        //            dvData.RowFilter = "actiontype= 'BUILD'";
        //            DataTable dtBuild = dvData.ToTable();
        //            dbCol=dbCol.Substring(0, dbCol.IndexOf("flash_type") - 1);*/

        //        string dbRowCol = this.dbCol.Replace(dbFixedcol, "");
        //        string[] rowsName = dbRowCol.Substring(1, dbRowCol.Length - 1).Split(',');
        //        dbFixedcol = dbFixedcol + ",field_name,plan_type,plan_qty";

        //        for (int i = 0; i < dtBuild.Rows.Count; i++)
        //        {
        //            sqlSaveFixed = "  insert into m_serv_forecast( forecast_tab_name, " + dbFixedcol + ")  values('" + sheetName + "'";
        //            for (int j = 0; j < 4; j++)
        //            {
        //                sqlSaveFixed = sqlSaveFixed + ",";
        //                sqlSaveFixed = sqlSaveFixed + "'" + dtBuild.Rows[i][j].ToString() + "'";
        //            }
        //            //sqlSave = sqlSave + sqlSaveFixed;

        //            for (int rownum = 0; rownum < rowsName.Length; rownum++)
        //            {

        //                int ind = rowsName[rownum].IndexOf("_type");
        //                if (ind > 0)
        //                {
        //                    string v_field_name = rowsName[rownum].Replace("_type", "");
        //                    string v_plan_type = v_field_name + "_type";
        //                    string v_plan_QTY = v_field_name + "_qty";
        //                    DataTable dt = dsMappingTable.Tables[0];
        //                    DataRow[] drQTY = dt.Select("table_field_name='" + v_plan_QTY + "'");

        //                    if (dtBuild.Rows[i][drQTY[0][2].ToString()].ToString() != "")
        //                    {

        //                        DataRow[] drfield = dt.Select("table_field_name='" + v_plan_type + "'");

        //                        sqlSave = sqlSave + sqlSaveFixed;

        //                        sqlSave = sqlSave + ",";
        //                        sqlSave = sqlSave + "'" + v_plan_type + "'";


        //                        sqlSave = sqlSave + ",";
        //                        sqlSave = sqlSave + "'" + dtBuild.Rows[i][drfield[0][2].ToString()].ToString() + "'";

        //                        sqlSave = sqlSave + ",";
        //                        sqlSave = sqlSave + "'" + dtBuild.Rows[i][drQTY[0][2].ToString()].ToString() + "'";
        //                        sqlSave = sqlSave + ")";
        //                    }
        //                }
        //            }

        //        }
        //        sqlSave = sqlDelete + sqlSave;
        //        SqlConnection cn = new SqlConnection(constr);
        //        try
        //        {

        //            cn.Open();
        //            SqlCommand cmd = new SqlCommand(sqlSave, cn);
        //            cmd.CommandTimeout = 0;
        //            cmd.ExecuteNonQuery();
        //            cn.Close();
        //            MessageBox.Show("Firm Demand has been uploaded successfully");
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("There is some error to save. Please check." + ex.Message.ToString());
        //            cn.Close();
        //        }
        //    }
        //    else
        //    {
        //        DataTable dtBuild = dataGridView1.DataSource as DataTable;
        //        string sqlDelete = " delete from m_serv_forecast where status='FLEXIBILITY BUFFER' ";
        //        String sqlSave = "";
        //        for (int i = 0; i < dtBuild.Rows.Count; i++)
        //        {

        //            sqlSave = sqlSave + "  insert into m_serv_forecast( forecast_tab_name,status,plan_type, plan_qty, delivery_date,forecast_org, field_name )  values('" + sheetName + "','FLEXIBILITY BUFFER'";

        //            sqlSave = sqlSave + ",";
        //            sqlSave = sqlSave + "'" + dtBuild.Rows[i][0].ToString() + "','" + dtBuild.Rows[i][1].ToString() + "','" + dtBuild.Rows[i][2].ToString() + "','" + dtBuild.Rows[i][5].ToString() + "','" + getAssemblyField(dtBuild.Rows[i][0].ToString()) + "'";

        //            sqlSave = sqlSave + ")";
        //        }


        //        sqlSave = sqlDelete + sqlSave;

        //        SqlConnection cn = new SqlConnection(constr);
        //        try
        //        {

        //            cn.Open();
        //            SqlCommand cmd = new SqlCommand(sqlSave, cn);
        //            cmd.CommandTimeout = 0;
        //            cmd.ExecuteNonQuery();
        //            cn.Close();
        //            MessageBox.Show("Flexibility buffer has been uploaded successfully");
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("There is some error to save. Please check." + ex.Message.ToString());
        //            cn.Close();
        //        }
        //    }
        //}
            //    MessageBox.Show(sql);
    
                    


        public string getAssemblyField(string planning_item)
        {
            string fld_name = "";
            SqlConnection cn = new SqlConnection(constr);
            try
            {               
                cn.Open();
                SqlCommand cmd = new SqlCommand("SELECT field_name FROM  m_serv_forecast_asm_mapping where planning_item='"+planning_item+"'", cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                cmd.CommandTimeout = 0;
                DataSet ds = new DataSet();
                da.Fill(ds);
               
                if (ds != null)
                    if (ds.Tables.Count>0)
                        if (ds.Tables[0].Rows.Count>0)
                                fld_name=ds.Tables[0].Rows[0][0].ToString();               

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error to get assembly fieldname. Please check." + ex.Message.ToString());

            }
            finally
            {
                cn.Close();
            }
            return fld_name;
        }



        //private void btnDelete_Click(object sender, EventArgs e)
        //{
        //    for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
        //    {
        //        dataGridView1.Rows.RemoveAt(dataGridView1.SelectedCells[i].RowIndex);
        //    }
        //}

        private void lstWorksheet_SelectedIndexChanged(object sender, EventArgs e)
        {
          /*  if (lstWorksheet.SelectedItems.Count > 0)
              {
                  ListItem1 lst = lstWorksheet.SelectedItem as ListItem1;               
                  ListQuery query = new ListQuery(lst.ID.ToString());
                  sheetName=lst.Name.ToString().Substring(0,5);
              //    SetListListView(this.spreadsheetService.Query(query));
              }*/ 
        }
        string WorkSheet = "";
        //private void button1_Click(object sender, EventArgs e)
        //{
        //    if (lstWorksheet.SelectedItems.Count > 0)
        //    {
        //        ListItem1 lst = lstWorksheet.SelectedItem as ListItem1;
        //        WorkSheet= lst.Name.ToString();
        //        ListQuery query = new ListQuery(lst.ID.ToString());
        //        //  sheetName = lst.Name.ToString().Substring(0,10);
        //        sheetName = ExtractNumbers(spreadsheetListView.SelectedItems[0].Text.Trim());
        //        SetListListView(this.spreadsheetService.Query(query));
        //    }
        //}

        static string ExtractNumbers(string expr)
        {
            return string.Join(null, System.Text.RegularExpressions.Regex.Split(expr, "[^\\d]"));
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            /* assembly mapping ***********/
         /*   if (tabControl1.SelectedTab == tbMappingTable)
            {
                loadPlanTypeFilter();
                dgAssmblyMapping.DataSource = null;
            }*/
        }


        /* ASSEMBLY MAPPING
            private void btnDeleteAss_Click(object sender, EventArgs e)
            {
                for (int i = 0; i < dgAssmblyMapping.SelectedRows.Count; i++)
                {
                    int id = Convert.ToInt32(dgAssmblyMapping.SelectedRows[i].Cells[0].Value.ToString());
                    string strdeleteAss = "delete from m_serv_forecast_asm_mapping where id=" + id + " select @@rowCount";
                    SqlConnection cn = new SqlConnection(constr);
                    cn.Open();
                    SqlCommand cmd = new SqlCommand(strdeleteAss, cn);
                    int rowAffected = cmd.ExecuteNonQuery();
                    if (rowAffected > 0)
                    {
                        MessageBox.Show(" Records have been deleted successfully ");
                    }
                }
                plan_type = cmbPlanType.Text;
                asm_heading = cmbPlanType.SelectedValue.ToString();
                loadGrid(plan_type, asm_heading);


            }
        */

            string plan_type = "";
            string asm_heading = "";
       /******Aseembly mapping    private void btnSaveAss_Click(object sender, EventArgs e)
            {
                DataTable dtAss = this.dgAssmblyMapping.DataSource as DataTable;
                String strSaveAss = ""; //"  delete from m_serv_forecast_asm_mapping  where plan_type='" + plan_type + "' " + whereclause;
                string[] columnName = asm_heading.Split(',');
                for (int i = 0; i < dtAss.Rows.Count; i++)
                {

                    String columnVal = "";
                    for (int j = 1; j < dtAss.Columns.Count; j++)
                    {
                        if (j == 1)
                            columnVal = columnVal + "'";
                        columnVal = columnVal + dtAss.Rows[i][j].ToString() + "','";
                    }
                    if (dtAss.Rows[i][0].ToString() == "")
                        strSaveAss = strSaveAss + "  insert into m_serv_forecast_asm_mapping  (plan_type," + asm_heading + ") values (" + columnVal.Substring(0, columnVal.Length - 2) + ")    ";
                    else
                    {
                        strSaveAss = strSaveAss + " update m_serv_forecast_asm_mapping   set plan_type='" + dtAss.Rows[i][1].ToString() + "'";
                        for (int cnt = 0; cnt < columnName.Length; cnt++)
                        {
                            int gridcolumncnt = cnt + 2;
                            strSaveAss = strSaveAss + "," + columnName[cnt] + "='" + dtAss.Rows[i][gridcolumncnt].ToString() + "'";
                        }
                        strSaveAss = strSaveAss + "  where id=" + dtAss.Rows[i][0].ToString();
                    }
                    //strSaveAss = strSaveAss + "update m_serv_forecast_asm_mapping  set plan_type='"++ "'"  //"      insert into m_serv_forecast_asm_mapping  (plan_type," + asm_heading + ") values (" + columnVal.Substring(0, columnVal.Length - 2) + ")    ";
                }
                //MessageBox.Show(strSaveAss);
                try
                {
                    SqlConnection cn = new SqlConnection(constr);
                    cn.Open();
                    SqlCommand cmd = new SqlCommand(strSaveAss, cn);
                    int rowAffected = cmd.ExecuteNonQuery();
                    MessageBox.Show(" Records have been updated successfully ");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }***/
            bool loaded = false;


            private void existPlanType()
            {

            }


            /* ASSEMBLY MAPPING  public void loadPlanTypeFilter()
              {
                  string sqlPlan = "select plan_type,asm_map_heading from m_serv_forecast_plantype_master ";
                  DataSet dsPlanType = getDataSet(sqlPlan);
                  DataRow dr = dsPlanType.Tables[0].NewRow();
                  dr["plan_type"] = "Select";
                  dr["asm_map_heading"] = "";
                  dsPlanType.Tables[0].Rows.InsertAt(dr, 0);
                  cmbPlanType.DataSource = dsPlanType.Tables[0];
                  cmbPlanType.DisplayMember = "plan_type";
                  cmbPlanType.ValueMember = "asm_map_heading";
                  loaded = true;
              }*/

          /*   public void loadCPUTypeFilter()
             {
                 string sqlPlan = "select distinct cpu_type dbo.m_serv_forecast_asm_mapping  ";
                 DataSet dsPlanType = getDataSet(sqlPlan);
                 cmbPlanType.DataSource = dsPlanType.Tables[0];
                 cmbPlanType.DisplayMember = "plan_type";
                 cmbPlanType.ValueMember = "asm_map_heading";

             }
             public void loadItrayTypeFilter()
             {
                 string sqlPlan = "select plan_type,asm_map_heading from m_serv_forecast_plantype_master ";
                 DataSet dsPlanType = getDataSet(sqlPlan);
                 cmbPlanType.DataSource = dsPlanType.Tables[0];
                 cmbPlanType.DisplayMember = "plan_type";
                 cmbPlanType.ValueMember = "asm_map_heading";

             }
             public void loadItrayDiskTypeFilter()
             {
                 string sqlPlan = "select plan_type,asm_map_heading from m_serv_forecast_plantype_master ";
                 DataSet dsPlanType = getDataSet(sqlPlan);
                 cmbPlanType.DataSource = dsPlanType.Tables[0];
                 cmbPlanType.DisplayMember = "plan_type";
                 cmbPlanType.ValueMember = "asm_map_heading";

             }
             public void loadRackTypeFilter()
             {
                 string sqlPlan = "select plan_type,asm_map_heading from m_serv_forecast_plantype_master ";
                 DataSet dsPlanType = getDataSet(sqlPlan);
                 cmbPlanType.DataSource = dsPlanType.Tables[0];
                 cmbPlanType.DisplayMember = "plan_type";
                 cmbPlanType.ValueMember = "asm_map_heading";

             } */


        DataSet dsAssembly;
        string whereclause = "";

        /************assembly mapping ******/
        //public void loadGrid(string plan_type, string asm_heading)
        //{
        //    String[] col = asm_heading.Split(',');

        //    String qryPart = "";
        //    //  int x1 = 30;
        //    //   int y1 =38;
        //    whereclause = "";
        //    for (int i = 0; i < col.Count(); i++)
        //    {
        //        /*  if (! col[i].Contains("ssd"))
        //          {
        //              // {
        //              x1 = x1 + 300;
                  
        //              // }
        //              //else
        //              //{
        //              //    x1=x1;
        //              //    y1=y1+50;
        //              //}
        //            //  getFilter(col[i], x1, y1);
        //          }*/

        //        if (qryPart != "")
        //            qryPart = qryPart + ", ";
        //        qryPart = qryPart + col[i] + " as [" + col[i].Replace("_", " ") + "]";
        //        try
        //        {
        //            ComboBox cmb = (ComboBox)(panel5.Controls.Find("cmb" + col[i], true)[0]);
        //            if (cmb.SelectedValue != "ALL")

        //                whereclause = whereclause + " and " + col[i] + " = '" + cmb.SelectedValue + "'";
        //        }
        //        catch
        //        { }


        //    }

        //    if (asm_heading != "")
        //    {

        //        string sqlAssembly = "select id, plan_type as [PLAN TYPE]," + qryPart.ToUpper() + " from m_serv_forecast_asm_mapping    where plan_type='" + plan_type + "'  " + whereclause;

        //        dsAssembly = getDataSet(sqlAssembly);
        //        this.dgAssmblyMapping.DataSource = null;
        //        this.dgAssmblyMapping.DataSource = dsAssembly.Tables[0];
        //        dgAssmblyMapping.Columns[0].Visible = false;

        //        for (int k = 0; k < dgAssmblyMapping.ColumnCount; k++)
        //        {
        //            dgAssmblyMapping.Columns[k].Width = 200;
        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show("There is no data for this plan_type");
        //    }
        //}


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

     /* ASSEMBLY MAPPING
      * 
      * private void btnLoadAssemblyGrid_Click(object sender, EventArgs e)
        {
            plan_type = cmbPlanType.Text;
            asm_heading = cmbPlanType.SelectedValue.ToString();
            loadGrid(plan_type, asm_heading);
        }*/

        //private void spreadsheetListView_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (spreadsheetListView.SelectedItems.Count > 0)
        //    {
        //        // Get the worksheet feed from the selected entry
        //        // MessageBox.Show(spreadsheetListView.SelectedItems[0].SubItems[1].Text);
        //        WorksheetQuery query = new WorksheetQuery(spreadsheetListView.SelectedItems[0].SubItems[1].Text);

        //        SetWorksheetListView(spreadsheetService.Query(query));
        //    }
        //}

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        int selectedRowIndex = 0;
        object[] Datarow1;

        private void dgAssmblyMapping_KeyDown(object sender, KeyEventArgs e)
        {



        }

   /* ASSEMBLY MAPPING     private void btnPasteRow_Click(object sender, EventArgs e)
        {

            try
            {

                for (int i = 0; i < dgAssmblyMapping.SelectedRows.Count; i++)
                {
                    DataRow dr = dsAssembly.Tables[0].NewRow();
                    for (int j = 1; j < this.dsAssembly.Tables[0].Columns.Count; j++)
                    {
                        dr[j] = this.dsAssembly.Tables[0].Rows[dgAssmblyMapping.SelectedRows[i].Index][j];

                    }

                    dsAssembly.Tables[0].Rows.Add(dr);
                }
                dgAssmblyMapping.DataSource = dsAssembly.Tables[0];
                MessageBox.Show(dgAssmblyMapping.SelectedRows.Count + " rows have been inserted  successfully");
            }
            catch
            {
            }

        }*/

        /* ASSEMBLY MAPPING   private void cmbPlanType_SelectedIndexChanged(object sender, EventArgs e)
           {
               if (loaded)
               {
                   panel5.Controls.Clear();
                   loadFilter();
               }

           }

           private void loadFilter()
           {
               plan_type = cmbPlanType.Text;
               asm_heading = cmbPlanType.SelectedValue.ToString();
               if (asm_heading != "")
               {
                   String[] col = asm_heading.Split(',');

                   String qryPart = "";
                   int x1 = 50;
                   int y1 = 38;
                   for (int i = 0; i < col.Count(); i++)
                   {
                       if (!col[i].Contains("ssd"))
                       {
                           if (i != 0)
                           {
                               int spbetlbtext = 20;
                               int lblsize = 120;
                               int cmbsize = 150;

                               x1 = x1 + lblsize + cmbsize + spbetlbtext;
                           }
                           getFilter(col[i], x1, y1);
                       }
                   }
               }
           }*/

        int strlen = 0;
        //private void getFilter(string colName, int xLoc, int yLoc)
        //{

        //    string strqry = " select distinct cast(" + colName + " as varchar) as " + colName + " from m_serv_forecast_asm_mapping    where " + colName + " is not null or " + colName + " <>''   order by " + colName;
        //    string dsname = "ds" + colName;

        //    DataSet dscombo = getDataSet(strqry);
        //    Label lbl = new Label();
        //    string colNameLabel = "";
        //    if (colName.Length > 10)
        //        colNameLabel = colName.Replace("_", " ").Substring(0, 12).ToUpper();
        //    else
        //        colNameLabel = colName.Replace("_", " ").ToUpper();

        //    lbl.Text = colNameLabel.ToUpper();
        //    lbl.Name = lbl + colName;
        //    lbl.Location = new Point(xLoc, yLoc);

        //    lbl.AutoSize = true;
        //    lbl.Font = new System.Drawing.Font("Verdana", 8.25F,System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        //    lbl.ForeColor = System.Drawing.SystemColors.ControlLightLight;
        //    lbl.Size = new System.Drawing.Size(120, 16);
        //    lbl.TextAlign = ContentAlignment.MiddleRight;


        //    ComboBox cmb = new ComboBox();
        //    cmb.Name = "cmb" + colName;
        //    cmb.Size = new System.Drawing.Size(150, 22);
        //    strlen = 100;
        //    int spbetlbtext = 20;
        //    cmb.Location = new Point(xLoc + strlen + spbetlbtext, yLoc);
        //    DataRow dr = dscombo.Tables[0].NewRow();
        //    dr[colName] = "ALL";
        //    dscombo.Tables[0].Rows.InsertAt(dr, 0);
        //    cmb.DataSource = dscombo.Tables[0];

        //    cmb.DisplayMember = colName;
        //    cmb.ValueMember = colName;


        //  //  panel5.Controls.Add(lbl);
        // //   panel5.Controls.Add(cmb);

        //}

        private void dgAssmblyMapping_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


        }


        public class ListItem1
        {
            private string id = string.Empty;
            private string name = string.Empty;
            public ListItem1(string sid, string sname)
            {
                id = sid;
                name = sname;
            }

            public override string ToString()
            {
                return this.name;
            }

            public string ID
            {
                get
                {
                    return this.id;
                }
                set
                {
                    this.id = value;
                }
            }

            public string Name
            {
                get
                {
                    return this.name;
                }
                set
                {
                    this.name = value;
                }
            }
        }

        //private void button11_Click(object sender, EventArgs e)
        //{
        //    SqlConnection cn = new SqlConnection(constr);
        //    //string strsql = "SELECT   forecast_tab_name , uPorID, Actionid, action_type, forecast_org, delivery_date, cpu_type, itray_type, itray_qty, itray_disk_type, itray_disk_qty,                     itray_ram_type, itray_ram_qty, itray_ssd_size, itray_ssd_type, itray_ssd_class, itray_ssd_qty, stray_type, stray_qty, rack_type, rack_vanilla, rack_qty, rack_switch_type, rack_switch_qty_per, status From  m_serv_forecast ";

        //    string strsql = "select forecast_tab_name,status,uporID,forecast_org,delivery_date,field_name,plan_type,plan_qty  From  m_serv_forecast ";
        //    SqlCommand cmd = new SqlCommand(strsql, cn);
        //    SqlDataAdapter da = new SqlDataAdapter(cmd);
        //    DataSet ds = new DataSet();
        //    da.Fill(ds);
        //    dgDownload.DataSource = ds.Tables[0];
        //    dgDownload.Columns[0].ReadOnly = true;
        //    dgDownload.Columns[dgDownload.ColumnCount - 1].Width = 200;
        //}

        //private void button10_Click(object sender, EventArgs e)
        //{
        //    for (int i = 0; i < dgDownload.SelectedRows.Count; i++)
        //    {
        //        this.dgDownload.Rows.RemoveAt(dgDownload.SelectedCells[i].RowIndex);
        //    }
        //}

        //private void btnSaveData_Click(object sender, EventArgs e)
        //{
        //    DataTable dt = dgDownload.DataSource as DataTable;
        //    string sqlSave = "  delete from m_serv_forecast  ";
           
        //    for (int i = 0; i < dt.Rows.Count; i++)
        //    {
        //        sqlSave = sqlSave + "  insert into m_serv_forecast ( forecast_tab_name,status,uporID,forecast_org,delivery_date,field_name,plan_type,plan_qty)  values(";
        //        for (int j = 0; j < dt.Columns.Count; j++)
        //        {
        //            sqlSave = sqlSave + "'" + dt.Rows[i][j].ToString() + "',"; ;
                   
        //        }
        //        sqlSave = sqlSave.Substring(0,sqlSave.Length-1) + ")";
        //    }

        //    try
        //    {
        //        SqlConnection cn = new SqlConnection(constr);
        //        cn.Open();
        //        SqlCommand cmd = new SqlCommand(sqlSave, cn);
        //        cmd.ExecuteNonQuery();
        //        cn.Close();
        //        MessageBox.Show("File has been uploaded successfully");
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("There is some error to save. Please check." + ex.Message.ToString());
        //    }
        //}

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void tbpageUpload_Click(object sender, EventArgs e)
        {

        }

        private void status_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void statusVersion_Click(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void lstWorksheet_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void worksheetListView_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tbMappingTable_Click(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnCopyrow_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel5_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void tbPageDownload_Click(object sender, EventArgs e)
        {

        }

        private void dgDownload_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
        string filename;
        string fname;
        private void spreadsheetListView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            
            if (spreadsheetListView1.SelectedItems.Count > 0)
            {
                filename = spreadsheetListView1.SelectedItems[0].SubItems[1].Text;
                fname = spreadsheetListView1.SelectedItems[0].SubItems[0].Text;
                //  googleshareDoc = spreadsheetListView.;
                //  fileId = spreadsheetListView.SelectedItems[0]..Text;
              //  Google.Apis.Drive.v2.Data.File file1 = service.Files.Get(filename).Execute();
                saveTo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\currentForecast" + System.DateTime.Now.Day.ToString("00") + System.DateTime.Now.Month.ToString("00")  + System.DateTime.Now.Year.ToString("00")  + System.DateTime.Now.Hour.ToString("00")  + System.DateTime.Now.Minute.ToString("00") + System.DateTime.Now.Second.ToString("00") + ".xlsx";
                saveTo = saveTo.Replace("file:\\", "");
                System.IO.FileStream fs = new System.IO.FileStream(saveTo, System.IO.FileMode.OpenOrCreate);
                fs.Close();

                if (downloadFile(service, filename, saveTo))
                {
                    lslWorksheet1.Items.Clear();
                    ListItem1 lst = new ListItem1("nsp_detailed_region", "nsp_detailed_region");
                    lslWorksheet1.Items.Add(lst);
                    ListItem1 lst1 = new ListItem1("deployed-raw", "deployed-raw");
                    lslWorksheet1.Items.Add(lst1);
                }                    
            }

            Cursor.Current = Cursors.Default;

        }
        Microsoft.Office.Interop.Excel.Worksheet sheet;
        Microsoft.Office.Interop.Excel.Sheets excelSheets;
       // Microsoft.Office.Interop.Excel.Application excelApp;
        Microsoft.Office.Interop.Excel.Workbook excelWorkbook=null;
        Microsoft.Office.Interop.Excel.Workbooks excelWorkbooks=null;
        int sheetnum = 0;
        public void getallSheet(string FileName)
        {          

           Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            try
            {
              // MessageBox.Show("getallsheet");              

                string workbookPath = FileName.Replace("file:\\", "");
                excelWorkbooks = excelApp.Workbooks;
                //MessageBox.Show("beforegetallsheet" + workbookPath);                
                excelWorkbook = excelWorkbooks.Open(workbookPath, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
             //   MessageBox.Show("geting all sheet");
                sheetnum = 0;
                excelSheets = excelWorkbook.Worksheets;
                lslWorksheet1.Items.Clear();
                foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in excelWorkbook.Worksheets)
                {
                        ListItem1 lst = new ListItem1(worksheet.Name, worksheet.Name);
                        lslWorksheet1.Items.Add(lst);                   
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
               // MessageBox.Show("file is loading.Please wait and try again in few second!","Message",MessageBoxButtons.OK);
            }
            finally
            {
                //Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
              /* excelWorkbook.Close(true, Type.Missing, Type.Missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);*/

               excelWorkbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
               excelWorkbooks.Close();
               excelApp.Quit();

                // System.Runtime.InteropServices.Marshal.ReleaseComObject(excelUsedRange);
              //  System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
              //  System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheets);
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
                // System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sheet);

                // System.Runtime.InteropServices.xlBook.Close(Type.Missing, Type.Missing, Type.Missing);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbooks);

                // excelApp.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);


            }



        }
      

        //public void showInGrid()
        //{



        //    System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + saveTo + "';Extended Properties=Excel 12.0;");
        //    System.Data.OleDb.OleDbDataAdapter MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [BP$]", MyConnection);
        //    MyCommand.TableMappings.Add("Table", "TestTable");
        //    System.Data.DataSet DtSet = new System.Data.DataSet();
        //    MyCommand.Fill(DtSet);
        //    DataTable myTable = DtSet.Tables[0];
        //    dataGridView1.DataSource = myTable;

        //    //   MyCommand.Fill(DtSet);
        //    //   dataGridView1.DataSource = DtSet.Tables[0];
        //    lblNumOfRow.Text = DtSet.Tables[0].Rows.Count.ToString();
        //    MyConnection.Close();


        //}

        private void button9_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (lslWorksheet1.SelectedItems.Count > 0)
            {
                ListItem1 lst = lslWorksheet1.SelectedItem as ListItem1;
                WorkSheet_name = lst.Name.ToString();
                foreach (Control ctrl in pnlPlanningItem.Controls)
                    pnlPlanningItem.Controls.Remove(ctrl);
                showInGrid(WorkSheet_name);
            }

            Cursor.Current = Cursors.Default;
        }
        int gListView1LostFocusItem;
        private void SpreadsheetListView_Leave(object sender, EventArgs e)
        {
            // Set the global int variable (gListView1LostFocusItem) to
            // the index of the selected item that just lost focus
            gListView1LostFocusItem = this.spreadsheetListView1.FocusedItem.Index;
        }

        private void listView1_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            // If this item is the selected item
            if (e.Item.Selected)
            {
                // If the selected item just lost the focus
                if (gListView1LostFocusItem == e.Item.Index)
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
                else if (spreadsheetListView1.Focused)  // If the selected item has focus
                {
                    // Set the colors to the normal colors for a selected item
                    e.Item.ForeColor = SystemColors.HighlightText;
                    e.Item.BackColor = SystemColors.Highlight;
                }
            }
            else
            {
                // Set the normal colors for items that are not selected
                e.Item.ForeColor = spreadsheetListView1.ForeColor;
                e.Item.BackColor = spreadsheetListView1.BackColor;
            }
            e.DrawBackground();
            e.DrawText();
        }


        string columnName = "";
        private void getColumnName(string tablename)
        {
            System.Data.DataSet ds = new System.Data.DataSet();
            string sqlQry = "select  *  from m_nsp_header_mapping where table_name like '" + tablename + "%'";
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
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                columnName = columnName + dr["table_field_name"] + ",";
            }
            columnName = columnName.Substring(0, columnName.Length - 1);
        }


        private bool checkExistInMapping()
        {
            DataTable dtBuild = this.dgvNewforecast.DataSource as DataTable;
            DataTable dtDistinct = dtBuild.DefaultView.ToTable(true, "planning_item");
           // dataTable.DefaultView.ToTable(true, "employeeid");
            System.Data.DataSet ds = new System.Data.DataSet();
            string sqlQry = "select distinct planning_item  from m_nsp_asm_mapping ";
            try
            {
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlCommand cmd = new SqlCommand(sqlQry, cn);
                cmd.CommandTimeout = 0;
                SqlDataAdapter da=new SqlDataAdapter(cmd);
                ds = new System.Data.DataSet();
                da.Fill(ds);
                cn.Close();
              //MessageBox.Show("File has been uploaded successfully");
            }
            catch (Exception ex)
            {
                //MessageBox.Show("There is some error to save. Please check." + ex.Message.ToString());
            }
            int y=0;
            pnlPlanningItem.AutoScroll = true;
            foreach (DataRow dr in dtDistinct.Rows)
            {
              if (dr[0].ToString()!="")
              if (ds.Tables[0].Select(" planning_item='" + dr[0].ToString() + "'").Length == 0)
              {
                  TextBox lbl=new TextBox();
                  lbl.Name=dr[0].ToString();
                  lbl.Text=dr[0].ToString();
                  lbl .Width = 533;
                  lbl.Height = 20;
                  lbl.ForeColor = System.Drawing.Color.Black;
                  lbl.Location = new Point(0, y);
                  lbl.ReadOnly = true;
               // lbl.BorderStyle = BorderStyle.None;
                  pnlPlanningItem.Controls.Add(lbl); 
                  y=y+20;
              }
            }
            if (pnlPlanningItem.Controls.Count>0)
                return false;
            else
                return true;

        }

        private void btnUploadNewForecast_Click(object sender, EventArgs e)
        {
              string tablename = "m_nsp_plan";
                if (radioButton2.Checked)
                    tablename = "m_nsp_compare_plan";
                if (radioButton3.Checked)
                    tablename = "m_nsp_deployed_raw";
              // DataTable dtBuild = this.dgvNewforecast.DataSource as DataTable;

                string sheetName = lslWorksheet1.SelectedItems[0].ToString();
            if  (( !radioButton3.Checked) && (!checkExistInMapping()))
            {    
           
                lblMappingTable.Visible = true;
                pnlPlanningItem.Visible = true;
                MessageBox.Show("Parts are not in the mapping table.NSP Plan can not be uploaded!");
             }
            else
            {
                pnlPlanningItem.Visible = false;
                lblMappingTable.Visible = false;
              //  MessageBox.Show(dgvNewforecast.Rows[0].Cells[6].Value.ToString());
           //     getColumnName(tablename);
             //   getExcelcolumnNameNSP(tablename);
             

                DataTable dtBuild = new DataTable();
               dtBuild = this.dgvNewforecast.DataSource as DataTable;
                string sqlDelete = "  delete from " + tablename;
                string sqlSave = "";
                string sqlSaveFixed = "";
                string[] colname = dbCol.Split(',');
                string[] datetype = dbDateType.Split(',');
              
                for (int i = 0; i < dtBuild.Rows.Count; i++)
                {
                    if (dtBuild.Rows[i][2].ToString().Trim() != "")
                    {
                        sqlSaveFixed = sqlSaveFixed + "  insert into " + tablename + "( forecast_tab_name, " + dbCol + /* forecast_tab_name,region,category,planning_item,delivery_date,qty*/ " )  values ('" + sheetName + "'";
                        for (int j = 0; j < colname.Length; j++)
                        {
                            sqlSaveFixed = sqlSaveFixed + ",";
                            if (datetype[j] == "int")
                            //   if (j == colname.Length - 1)
                            {
                                //sqlSaveFixed = sqlSaveFixed + "'" + dtBuild.Rows[i][colname[j]].ToString().Replace(",", "") + "'";

                                int qty = 0;
                                if (dgvNewforecast.Rows[i].Cells[colname[j]].Value.ToString() != "")
                                {
                                    decimal decQty = Convert.ToDecimal(dgvNewforecast.Rows[i].Cells[colname[j]].Value.ToString().Replace(",", ""));
                                    qty = Decimal.ToInt32(decQty);
                                }
                                sqlSaveFixed = sqlSaveFixed + "'" + qty + "'";
                            }
                            else
                            {
                                sqlSaveFixed = sqlSaveFixed + "'" + dgvNewforecast.Rows[i].Cells[colname[j]].Value.ToString() + "'";// dtBuild.Rows[i][colname[j]].ToString() + "'";
                            }
                        }
                        sqlSaveFixed = sqlSaveFixed + ")";
                    }
                }

                sqlSave = sqlDelete + sqlSaveFixed;
                string sqlupdate = "update nsp set field_name=m_nsp_asm_mapping.field_name,plan_group=m_nsp_asm_mapping.plan_group  from m_nsp_asm_mapping inner join " + tablename + " as nsp on m_nsp_asm_mapping.planning_item=nsp.planning_item ";
                string sqlcreatestep = " update MRP_PROCESS_STEP_TABLE set m_mrp_Plan_update_time=GETDATE()  where m_mrp_plan_name='NSP Plan' ";
            if  ( !radioButton3.Checked)
                sqlSave = sqlSave + sqlupdate + sqlcreatestep;

            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlTransaction tran=cn.BeginTransaction();
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
           }
        }

        

        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
             for (int i = 0; i < this.dgvNewforecast.SelectedRows.Count; i++)
             {
                 dgvNewforecast.Rows.RemoveAt(dgvNewforecast.SelectedCells[i].RowIndex);
             }
        }

        private void button12_Click(object sender, EventArgs e)
        {

        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
          /* Microsoft.SqlServer.Dts.Runtime.Application   app =new Microsoft.SqlServer.Dts.Runtime.Application();
           string file = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\nsp_demand.dtsx";
           MessageBox.Show(file);
            Package package = app.LoadPackage(file, null);
            //'Dim vars As Variables = package.Variables
            //'vars("filename").Value = "dtsfilename"
             //'vars("filePath").Value = ""
            try{         
                DTSExecResult res  = package.Execute();
                if (res== Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure) 
                    { m
                        foreach (Microsoft.SqlServer.Dts.Runtime.DtsError local_DtsError in  package.Errors) 
                        { 
                            MessageBox.Show(local_DtsError.Description.ToString()); 
                        } 
                    } */

            Cursor.Current=Cursors.WaitCursor;
            SqlConnection cn = new SqlConnection(constrMVReports);
            string sqlStatus=" exec [dbo].[m_executeNSPJob] 'nsp_demand_supply_usage'";
             try
            {
                
                cn.Open();
                SqlCommand cmd = new SqlCommand(sqlStatus, cn);
                cmd.CommandTimeout = 0;
                SqlDataReader dr= cmd.ExecuteReader();
                 if (dr.Read())
                                
                 if (dr[0].ToString()=="Idle.")
                     MessageBox.Show("Job is successfully completed");
                 cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error to save. Please check." + ex.Message.ToString());
                cn.Close();
            }
            
             Cursor.Current=Cursors.Default;
  }
        string WorkSheet_name = "";
        private void lslWorksheet1_SelectedIndexChanged(object sender, EventArgs e)
        {

          
        } 
    
    
        //private void btnArchiveData_Click(object sender, EventArgs e)
        //{
        //    if (WorkSheet.Contains("Firm Demand"))
        //    {
        //        DataTable dtBuild = dataGridView1.DataSource as DataTable;
        //     //   string sqlDelete = "delete from m_serv_forecast where status<>'FLEXIBILITY BUFFER' ";
        //        string sqlSave = "";
        //        string sqlSaveFixed = "";
        //        /*    DataView dvData = new DataView(dt);
        //            dvData.RowFilter = "actiontype= 'BUILD'";
        //            DataTable dtBuild = dvData.ToTable();
        //            dbCol=dbCol.Substring(0, dbCol.IndexOf("flash_type") - 1);*/

        //        string dbRowCol = this.dbCol.Replace(dbFixedcol, "");
        //        string[] rowsName = dbRowCol.Substring(1, dbRowCol.Length - 1).Split(',');
        //        dbFixedcol = dbFixedcol + ",field_name,plan_type,plan_qty";

        //        for (int i = 0; i < dtBuild.Rows.Count; i++)
        //        {
        //            sqlSaveFixed = "  insert into m_serv_forecast_archive( forecast_tab_name, " + dbFixedcol + ")  values('" + sheetName + "'";
        //            for (int j = 0; j < 4; j++)
        //            {
        //                sqlSaveFixed = sqlSaveFixed + ",";
        //                sqlSaveFixed = sqlSaveFixed + "'" + dtBuild.Rows[i][j].ToString() + "'";
        //            }
        //            //sqlSave = sqlSave + sqlSaveFixed;

        //            for (int rownum = 0; rownum < rowsName.Length; rownum++)
        //            {

        //                int ind = rowsName[rownum].IndexOf("_type");
        //                if (ind > 0)
        //                {
        //                    string v_field_name = rowsName[rownum].Replace("_type", "");
        //                    string v_plan_type = v_field_name + "_type";
        //                    string v_plan_QTY = v_field_name + "_qty";
        //                    DataTable dt = dsMappingTable.Tables[0];
        //                    DataRow[] drQTY = dt.Select("table_field_name='" + v_plan_QTY + "'");

        //                    if (dtBuild.Rows[i][drQTY[0][2].ToString()].ToString() != "")
        //                    {

        //                        DataRow[] drfield = dt.Select("table_field_name='" + v_plan_type + "'");

        //                        sqlSave = sqlSave + sqlSaveFixed;

        //                        sqlSave = sqlSave + ",";
        //                        sqlSave = sqlSave + "'" + v_plan_type + "'";


        //                        sqlSave = sqlSave + ",";
        //                        sqlSave = sqlSave + "'" + dtBuild.Rows[i][drfield[0][2].ToString()].ToString() + "'";

        //                        sqlSave = sqlSave + ",";
        //                        sqlSave = sqlSave + "'" + dtBuild.Rows[i][drQTY[0][2].ToString()].ToString() + "'";
        //                        sqlSave = sqlSave + ")";
        //                    }
        //                }
        //            }

        //        }
        //       // sqlSave = sqlDelete + sqlSave;
        //        SqlConnection cn = new SqlConnection(constrMVAPPS);
        //        try
        //        {

        //            cn.Open();
        //            SqlCommand cmd = new SqlCommand(sqlSave, cn);
        //            cmd.CommandTimeout = 0;
        //            cmd.ExecuteNonQuery();
        //            cn.Close();
        //            MessageBox.Show("Firm Demand has been archived successfully");
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("There is some error to save. Please check." + ex.Message.ToString());
        //            cn.Close();
        //        }
        //    }
        //  else
        //    {
        //        DataTable dtBuild = dataGridView1.DataSource as DataTable;
        //        DataSet dsMappingColumn=getDataSet("SELECT  *  FROM  m_serv_flex_header_mapping ");
        //        string sharedDocColName="";
               
        //        //string sqlDelete = " delete from m_serv_forecast where status='FLEXIBILITY BUFFER' ";
        //        String sqlSave = "";

        //        DataTable dtColumn=new DataTable();
        //        dtColumn.Columns.Add(new DataColumn("tab_col",typeof(string)));
        //        dtColumn.Columns.Add(new DataColumn("shared_doc_col",typeof(string)));

        //            for (int j=0;j<dsMappingColumn.Tables[0].Rows.Count;j++)                     

        //                for(int col=0;col<dtBuild.Columns.Count;col++)
        //        {
        //            string tbl_shared=dsMappingColumn.Tables[0].Rows[j][2].ToString();
        //            string shared_tbl = dtBuild.Columns[col].ColumnName.ToString();
                  
        //          if ( dsMappingColumn.Tables[0].Rows[j][2].ToString().Contains(dtBuild.Columns[col].ColumnName.ToString()))
        //          {
        //              DataRow dr = dtColumn.NewRow();
        //              dr[0] = dsMappingColumn.Tables[0].Rows[j][1].ToString();
        //              dr[1]=dtBuild.Columns[col].ColumnName.ToString();
        //              dtColumn.Rows.Add(dr);   

        //        }
                     
        //            }


        //            for (int i = 0; i < dtBuild.Rows.Count; i++)
        //            {

        //                sqlSave = sqlSave + "  insert into m_serv_forecast_archive( forecast_tab_name,status,plan_type, plan_qty, delivery_date,forecast_org, field_name )  values('" + sheetName + "','FLEXIBILITY BUFFER'";

        //                sqlSave = sqlSave + ",'";

        //                DataRow[] drplanType=dtColumn.Select("tab_col='plan_type'");
        //                if (drplanType.Length > 0)
        //                    sqlSave = sqlSave + dtBuild.Rows[i][drplanType[0][1].ToString()].ToString() + "','";
        //                else

        //                    sqlSave = sqlSave + "','";
        //                DataRow[] drPlanQty = dtColumn.Select("tab_col='plan_qty'");
        //                if (drPlanQty.Length > 0)
        //                    sqlSave = sqlSave + dtBuild.Rows[i][drPlanQty[0][1].ToString()].ToString() + "','";
        //                else

        //                    sqlSave = sqlSave + "','";
        //                DataRow[] drDelDate = dtColumn.Select("tab_col='delivery_date'");
        //                if (drDelDate.Length > 0)
        //                    sqlSave = sqlSave + dtBuild.Rows[i][drDelDate[0][1].ToString()].ToString() + "','";
        //                else

        //                    sqlSave = sqlSave + "','";
        //                DataRow[] drOrg = dtColumn.Select("tab_col='forecast_org'");
        //                if (drOrg.Length > 0)
        //                    sqlSave = sqlSave + dtBuild.Rows[i][drOrg[0][1].ToString()].ToString() + "','";
        //                else
        //                    sqlSave = sqlSave + "','";
        //                sqlSave = sqlSave + getAssemblyField(dtBuild.Rows[i][drplanType[0][1].ToString()].ToString()) + "')";
                        
        //            }


        //       //// sqlSave = sqlDelete + sqlSave;

        //            SqlConnection cn = new SqlConnection(constrMVAPPS);
        //            try
        //            {

        //                cn.Open();
        //                SqlCommand cmd = new SqlCommand(sqlSave, cn);
        //                cmd.CommandTimeout = 0;
        //                cmd.ExecuteNonQuery();
        //                cn.Close();
        //                MessageBox.Show("Flexibility buffer has been archived successfully");
        //            }
        //            catch (Exception ex)
        //            {
        //                MessageBox.Show("There is some error to save. Please check." + ex.Message.ToString());
        //                cn.Close();
        //            }
        //    }
        //}

        //private static IAuthorizationState GetAuthorization(NativeApplicationClient client)
        //{
        //    // You should use a more secure way of storing the key here as
        //    // .NET applications can be disassembled using a reflection tool.
        //    const string STORAGE = "gdrive_uploader";
        //    const string KEY = "z},drdzf11x9;87"; //"AIzaSyDSdJ1EyURN8zPF20FwEfKPoeO-9YwqKYc";//
        //    string scope = Google.Apis.Drive.v2.DriveService.Scopes.Drive.GetStringValue();

        //    // Check if there is a cached refresh token available.
        //    IAuthorizationState state = AuthorizationMgr.GetCachedRefreshToken(STORAGE, KEY);
        //    if (state != null)
        //    {
        //        try
        //        {
        //            TimeSpan ts = new TimeSpan(0, 0, 30);
        //            client.RefreshToken(state, ts);
        //            return state; // Yes - we are done.
        //        }
        //        catch (DotNetOpenAuth.Messaging.ProtocolException ex)
        //        {
        //            Debug.WriteLine("Using existing refresh token failed: " + ex.Message);
        //        }
        //    }

        //    // If we get here, there is no stored token. Retrieve the authorization from the user.
        //    state = AuthorizationMgr.RequestNativeAuthorization(client, scope);
        //    AuthorizationMgr.SetCachedRefreshToken(STORAGE, KEY, state);
        //    return state;
        //}
        //private static IAuthenticator CreateAuthenticator()
        //{
        //    var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
        //    provider.ClientIdentifier = "679763664590-5tugcvfsp90pbc707t4f8dlma3oj9bab.apps.googleusercontent.com";// ClientCredentials.CLIENT_ID;
        //    provider.ClientSecret = "2lneARBZtpxrbX9vVkxcDiXe";// ClientCredentials.CLIENT_SECRET;
        //    IAuthorizationState st = GetAuthorization(provider);
        //    IAuthenticator auth = new OAuth2Authenticator<NativeApplicationClient>(provider, GetAuthorization);
        //    return auth;
        //}

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

    }
}
