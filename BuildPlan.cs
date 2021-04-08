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
using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Drive.v2.Data;
using Google.Apis.Services;


using Google.Apis.Drive.v2.Data;
using Google.Apis.Drive.v2;
using Google.Apis.Util;

using System.Net;

namespace PlanningExecution
{
    public partial class BuildPlan : Form
    {
      
        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
       string constrTEST ="";// ConfigurationManager.ConnectionStrings["epicorTESTConnectionString"].ConnectionString;
      //  string constr = "Data Source=Epicor;Initial Catalog=MIMDIST;user id=sa;password=mimi~100;Pooling=true;Max Pool Size=100;Min Pool Size=1;";

       string googleshareDoc;
        public BuildPlan()
        {
            loaded = false;          
            InitializeComponent();
            setSpreadsheetListView();
        }

    
           
        string filename;
        DriveService service;

        //get spreadsheet buildplan from google shared doc
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

                    FileList fileList = lstreq.Execute();
                    foreach (Google.Apis.Drive.v2.Data.File fileitem in fileList.Items)
                    {
                        if (fileitem.Title == "MIM Build Plans PROD")//"MIM NPI:" + projectName)
                        {   // titlestr = titlestr + "\n " + parentitem.Title;
                            Parentid = fileitem.Id; 
                            break;
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


            string file_id = getchildInFolder(service, Parentid);

            

           
            //for (int i = 0; i < entries.Count; i++)
            //{
            //    // Get the worksheets feed URI


            //    if (entries[i].Title.Text.StartsWith("Build Plan"))
            //    {
                  
            //        AtomLink worksheetsLink = entries[i].Links.FindService(GDataSpreadsheetsNameTable.WorksheetRel, AtomLink.ATOM_TYPE);
                   
            //        // Create an item that will display the title and hide the worksheets feed URI
                   
            //        ListViewItem item = new ListViewItem(new string[2] { entries[i].Title.Text, worksheetsLink.HRef.Content });
            //        this.spreadsheetListView.Items.Add(item);
            //    }
            //}
        }
        //get all worksheet for the selected buildplan put into worksheet Listview
  /* wwwwwwwwwwwwwwwww     private void worksheetListView_SelectedIndexChanged(object sender, EventArgs e)
        {
           if (worksheetListView.SelectedItems.Count > 0)
            {
                // MessageBox.Show(worksheetListView.SelectedItems[0].SubItems[2].Text);
                //ListQuery query = new ListQuery("https://docs.google.com/a/google.com/spreadsheets/d/111QGTrRrv6goVxJ1QsQI1qkg53Ao4kdaibyAoI2uq_g/edit#gid=2062635002", spreadsheetService);
                //worksheetListView.SelectedItems[0].SubItems[3].Text);
               ListQuery query = new ListQuery(worksheetListView.SelectedItems[0].SubItems[3].Text);
               SetListListView(this.spreadsheetService.Query(query));               
            }
        }wwwwwwwwwwwwwwwwww*/
        DataTable table = new DataTable();

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

        Thread t;
       
        private void button1_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (spreadsheetListView.SelectedItems.Count > 0)
            {

                filename = spreadsheetListView.SelectedItems[0].SubItems[1].Text;
                googleshareDoc = spreadsheetListView.SelectedItems[0].SubItems[0].Text;
                //  fileId = spreadsheetListView.SelectedItems[0]..Text;
                Google.Apis.Drive.v2.Data.File file1 = service.Files.Get(filename).Execute();
                saveTo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\currentBuildPaln.xlsx";
                saveTo = saveTo.Replace("file:\\", "");
                if (downloadFile(service, file1, saveTo))
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

            System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + saveTo + "';Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1\";");
            System.Data.OleDb.OleDbDataAdapter MyCommand = new System.Data.OleDb.OleDbDataAdapter("select F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15,F16,F17,F18,F19,F20,F21 from [BP$]  ", MyConnection);
            MyCommand.TableMappings.Add("Table", "TestTable");
            System.Data.DataSet DtSet = new System.Data.DataSet();
            MyCommand.Fill(DtSet);
            DataTable myTable = DtSet.Tables[0];
            myTable.Rows[0].Delete();
            myTable=  myTable.Select("F3 <>''").CopyToDataTable();
            dataGridView1.DataSource = myTable;
            dataGridView1.DisplayLayout.Bands[0].Columns["F1"].Width =80;
            dataGridView1.DisplayLayout.Bands[0].Columns["F2"].Width = 70;
            dataGridView1.DisplayLayout.Bands[0].Columns["F3"].Width = 100;
            dataGridView1.DisplayLayout.Bands[0].Columns["F4"].Width = 140;
            dataGridView1.DisplayLayout.Bands[0].Columns["F5"].Width = 30;
            dataGridView1.DisplayLayout.Bands[0].Columns["F6"].Width =80;
            dataGridView1.DisplayLayout.Bands[0].Columns["F7"].Width = 80;
            dataGridView1.DisplayLayout.Bands[0].Columns["F8"].Width = 350;
            dataGridView1.DisplayLayout.Bands[0].Columns["F9"].Width = 50;
            dataGridView1.DisplayLayout.Bands[0].Columns["F10"].Width = 80;
            dataGridView1.DisplayLayout.Bands[0].Columns["F11"].Width = 40;
            dataGridView1.DisplayLayout.Bands[0].Columns["F12"].Width = 100;
            dataGridView1.DisplayLayout.Bands[0].Columns["F13"].Width = 90;
            dataGridView1.DisplayLayout.Bands[0].Columns["F14"].Width = 50;
            dataGridView1.DisplayLayout.Bands[0].Columns["F15"].Width = 50;
            dataGridView1.DisplayLayout.Bands[0].Columns["F16"].Width =50;
            dataGridView1.DisplayLayout.Bands[0].Columns["F17"].Width = 50;
            dataGridView1.DisplayLayout.Bands[0].Columns["F18"].Width = 50;
            dataGridView1.DisplayLayout.Bands[0].Columns["F19"].Width = 120;
            dataGridView1.DisplayLayout.Bands[0].Columns["F20"].Width = 120;
            dataGridView1.DisplayLayout.Bands[0].Columns["F21"].Width = 80;
         //   MyCommand.Fill(DtSet);
      //     
         //   dataGridView1.DataSource = DtSet.Tables[0];
            lblNumOfRow.Text = myTable.Rows.Count.ToString();
            MyConnection.Close();
       

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
         getHeaderMapping();
            try
            {

               string sql = "exec m_process_build_plan '" + googleshareDoc + "','" + googleshareDoc.Replace("Build Plan ", "").Replace(".csv", "") + "'   exec m_process_build_plan_gig '" + googleshareDoc + "','" + googleshareDoc.Replace("Build Plan ", "").Replace(".csv", "") + "'";
               DataSet ds = getDataSet(sql);
               if (ds.Tables[0].Rows[0][0].ToString() == "0")
                   MessageBox.Show("BuildPlan has been  processed successfully");
               else
                   MessageBox.Show("Error to process BuildPlan. Please check!");

           }
           catch (Exception ex)
           {
               MessageBox.Show("Error to process BuildPlan. Please check!" + ex.Message.ToString());
           }
           Cursor.Current = Cursors.Default;       
      }

       DataSet dsMappingTable;
       private void getHeaderMapping()
       {
           string column_names="";
           string sqlstrHeaderMapping = "SELECT table_field_name,shared_doc_header FROM m_bp_header_mapping";
           SqlConnection cn = new SqlConnection(constr);
           cn.Open();
           SqlCommand cmd = new SqlCommand(sqlstrHeaderMapping, cn);
           SqlDataAdapter da = new SqlDataAdapter(cmd);
           dsMappingTable = new DataSet();
           da.Fill(dsMappingTable);
           table = dataGridView1.DataSource as DataTable;
           for (int i = 0; i < dsMappingTable.Tables[0].Rows.Count; i++)
           {
               column_names = column_names + dsMappingTable.Tables[0].Rows[i]["table_field_name"].ToString() + ",";
           }
           column_names = column_names.Substring(0,column_names.Length - 1);
           string sqlSaveFixed = "delete from m_temp_build_plan";
           for (int igd = 0; igd < table.Rows.Count; igd++)
           {
               if (table.Rows[igd]["F3"].ToString().Trim() != "build_id")
               {
                   sqlSaveFixed = sqlSaveFixed + "  insert into m_temp_build_plan( " + column_names + ")  values('";

                   for (int imap = 0; imap < dsMappingTable.Tables[0].Rows.Count; imap++)
                       sqlSaveFixed = sqlSaveFixed +table.Rows[igd][dsMappingTable.Tables[0].Rows[imap]["shared_doc_header"].ToString()].ToString().Replace("'","") + "','";
                   sqlSaveFixed = sqlSaveFixed.Substring(0, sqlSaveFixed.Length - 2) + ")";
               }
           }
           try
           {
               sqlSaveFixed = sqlSaveFixed + "   update m_temp_build_plan set DemandSourcetype=upper(DemandSourcetype) ";
                SqlConnection cnSave = new SqlConnection(constr);
               cnSave.Open();
               SqlCommand cmdSave = new SqlCommand(sqlSaveFixed, cnSave);
               cmdSave.ExecuteNonQuery();
           }
           catch (Exception e)
           {
               MessageBox.Show("Error is " + e.Message.ToString());
           }
          // SqlDataAdapter da = new SqlDataAdapter(cmd);
         //  dsMappingTable = new DataSet();
         //  da.Fill(dsMappingTable);

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

       private void button2_Click(object sender, EventArgs e)
       {

       }

       private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
       {

       }


       //private static IAuthenticator CreateAuthenticator()
       //{
       //    var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
       //    provider.ClientIdentifier = "679763664590-5tugcvfsp90pbc707t4f8dlma3oj9bab.apps.googleusercontent.com";// ClientCredentials.CLIENT_ID;
       //    provider.ClientSecret = "2lneARBZtpxrbX9vVkxcDiXe";// ClientCredentials.CLIENT_SECRET;
       //    IAuthorizationState st = GetAuthorization(provider);
       //    IAuthenticator auth = new OAuth2Authenticator<NativeApplicationClient>(provider, GetAuthorization);
       //    return auth;
       //}


       public static Google.Apis.Drive.v2.DriveService dr { get; private set; }
       static String CLIENT_ID = "943908261643.apps.googleusercontent.com";
       static String CLIENT_SECRET = "63oqKhgS7Gm9I3mx3PBVGcXn";
       static String REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob http://localhost";
       static String[] SCOPES = new String[] { "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/userinfo.email", "https://www.googleapis.com/auth/userinfo.profile" };

       private void button1_Click_1(object sender, EventArgs e)
       {
          // ListQuery query = new ListQuery("https://spreadsheets.google.com/feeds/cells/111QGTrRrv6goVxJ1QsQI1qkg53Ao4kdaibyAoI2uq_g/oy41ezk/private/full");
         //   SetListListView(this.spreadsheetService.Query(query));
         // Cursor.Current = Cursors.Hand;
           //if (lstWorksheet.SelectedItems.Count > 0)
           //{
           //    ListItem1 lst = lstWorksheet.SelectedItem as ListItem1;
           //    ListQuery query = new ListQuery(lst.ID.ToString());
           //    sheetName = lst.Name.ToString().Substring(0,10);
           //    sheetName = ExtractNumbers(spreadsheetListView.SelectedItems[0].Text.Trim());
           //    SetListListView(this.spreadsheetService.Query(query));
           //} 
                System.Data.OleDb.OleDbConnection MyConnection ;
                System.Data.DataSet DtSet ;
                System.Data.OleDb.OleDbDataAdapter MyCommand ;
                MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='e:\\ruchi\\testing.xlsx';Extended Properties=Excel 12.0;");
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [BP$]", MyConnection);
                MyCommand.TableMappings.Add("Table", "TestTable");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                dataGridView1.DataSource = DtSet.Tables[0];
                MyConnection.Close();    
       }
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

    
}
