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
using System.Data.OleDb;



namespace Version3
{
    public partial class NPI_MRP : Form
    {
        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        Microsoft.Office.Interop.Excel.Application excelApp1;
        Microsoft.Office.Interop.Excel.Worksheet sheet;
        Microsoft.Office.Interop.Excel.Workbook excelWorkbook;
        Microsoft.Office.Interop.Excel.Sheets excelSheets;
        string filename = "NPI_MRP";
        string projectName = "";
     
        public NPI_MRP()
        {
            InitializeComponent();
        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {

        }


        string build_id = "";
        string filterBuildID = "";
        private void btnDownloadGrid_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            string bd= String.Join(",", txtBuildID.Text.ToString().Split(',').Where((x, index) => index != 1).ToArray());
          // txtBuildID.Text = txtBuildID.Text.ToString().Replace(" ", "");
           build_id = txtBuildID.Text.ToString().Replace(" ", "").Replace(",", "','");
           if (radioButton1.Checked)
           {
               filterBuildID = " buildid in ('" + build_id + "')";
               filename = "NPI_MRP_" /*+ txtBuildID.Text.ToString() + "_" */+ DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Year.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + ".xlsx";
               filename = filename.Replace(',', '_').Replace(" 12:00:00 AM", "");
           }
           else
           {
               filterBuildID = " project_Name in ('" + txtProjectName.Text.Trim() + "')";
               filename = "NPI_MRP_"/* + txtProjectName.Text.ToString() + "_" */+ DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Year.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + ".xlsx";
               filename = filename.Replace(',', '_').Replace(" 12:00:00 AM", "");
               filefor = "Excel_doc";
           }
           //for (int i = 0; i < build_id.Count; i++)
           //{
           //    filterBuildID = "(" + filterBuildID + "'" + build_id[i] + "',";
           //}
            CreateExcel();
            cellUpdate();

            Cursor.Current = Cursors.Default;
        }

        public void copyPasteRangeExcel(string cellrng1, string cellrng2)
        {
            Microsoft.Office.Interop.Excel.Range sourceRange = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range("A17", "A17").EntireRow;
            Microsoft.Office.Interop.Excel.Range destinationRange = (Microsoft.Office.Interop.Excel.Range)sheet.get_Range(cellrng1, cellrng2);
            sourceRange.Copy(Type.Missing);
            //   destinationRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
         //   destinationRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType., Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            destinationRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            //excelApp.get_Range("A1:A360,B1:E1", Type.Missing).Merge(Type.Missing);            
            excelWorkbook.Save();
            // excelWorkbook.Close(true, false, Type.Missing);
        }

        public void CreateExcel()
        {         
             //  excelWorkbook.SaveAs("E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\"+filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();           
            //string workbookPath = "E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\VendorSheet.xls";
            if (filefor=="Excel_doc")
                filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + filename;
            else                           
                filename = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Templates\\" + filename ;
          //  MessageBox.Show(filename.Length.ToString());
           if (System.IO.File.Exists(filename))
               System.IO.File.Delete(filename);

          // string workbookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\Templates\\NPI_MRP_TEMPLATE.xlsx";
           string workbookPath ="\\\\mvfile\\MV REPORTS\\Master Data & Forms\\PlanningExecutionForms\\Templates\\NPI_MRP_TEMPLATE.xlsx";  
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook1 = excelApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
           
            try
            {
                excelWorkbook1.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            }
            catch (Exception e)
            {
                string msg = e.Message.ToString();
            }
            finally
            {
                //Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
                excelWorkbook1.Close(true, false, Type.Missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook1);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }

        }

        public int sheetnum=0;
        public void cellUpdate()
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            try
            {
                
                excelApp1 = new Microsoft.Office.Interop.Excel.Application();
                //string workbookPath = "E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\VendorSheet.xls"; Microsoft.Office.Interop.Excel.XlPlatform.xlWindows    
                string workbookPath =  filename;
           
                excelWorkbook = excelApp1.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                sheetnum = 0;
                excelSheets = excelWorkbook.Worksheets;
                sheetnum = sheetnum + 1;

             
                                string currentSheet = "TRAY-ASM";
                                string form_factor = "TRAY-ASM";
                                excelApp1.DisplayAlerts = false;
                                sheet = excelSheets[sheetnum] as Microsoft.Office.Interop.Excel.Worksheet;
                                updateBuildPlanDetail();
                                updateAssemblyDetail(form_factor);
                                //getdateHeader();

                                string strQuery = " exec sp_app_NPIMRP_compDetail_rev '" + txtBuildID.Text + "'";
                                DataSet dsBPDetail = getDataSet(strQuery);
                                if (dsBPDetail.Tables[1].Columns.Count > 1)
                                {
                                    updateComponentDetail_new(dsBPDetail.Tables[1]);
                                }
                
                                form_factor = "RACK-ASM";
                                excelApp1.DisplayAlerts = false;
                                sheet = excelSheets[sheetnum] as Microsoft.Office.Interop.Excel.Worksheet;
          
                                updateBuildPlanDetail();
                               updateAssemblyDetail(form_factor);
                              // updateComponentDetail_new(form_factor);
                               if (dsBPDetail.Tables[2].Columns.Count > 1)
                               {
                                   updateComponentDetail_new(dsBPDetail.Tables[2]);
                               }
                                excelApp1.DisplayAlerts = false;
                           /*     sheet = excelSheets[sheetnum] as Microsoft.Office.Interop.Excel.Worksheet;
                                updateCompPOInformation();*/
                                // getdateHeader();
                                // updateComponentDetail(form_factor);

                                sheet = excelSheets[sheetnum] as Microsoft.Office.Interop.Excel.Worksheet;

                                // updateCompPOInformation(sheet);
                                dumpPOInExcel(sheet);

                                excelWorkbook.Close(true, false, Type.Missing);

                                if (filefor != "Shared_doc")
                                {
                                    p.StartInfo.FileName = filename;
                                    //p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
                                    p.Start();
                                }
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp1);
                              
            }
            catch (Exception e)
            {             
                string Message = e.Message;
                MessageBox.Show(Message);
                if (excelApp1 != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp1);
            }
            finally
            {
                GC.Collect();
                p.Close();
            }
 
        }
        string concat_build_id="";
        public void updateBuildPlanDetail()
        {
            String strQuery = "SELECT [BP_BuildID]       ,[BP_ProjCode]      ,[BP_ProjName]      ,[BP_PopDate]      ,[BP_TTLAGPN]      ,[BP_TTLAQty]  FROM [MIMDIST].[dbo].[NPIMRP_BPDetails]  ";
            if (filterBuildID != "")
            {
                strQuery = strQuery + "  where BP_BuildID in ('" + build_id + "')";
            }
            DataSet dsBPDetail = getDataSet(strQuery + "  Order by BP_BuildID  ");
            int col = 0;
            concat_build_id = "";
            for (int i = 0; i < dsBPDetail.Tables[0].Rows.Count; i++)
            {
                //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());
                concat_build_id = concat_build_id + "," + dsBPDetail.Tables[0].Rows[i]["BP_BuildID"].ToString();
                for (int j = 0; j < dsBPDetail.Tables[0].Columns.Count; j++)
                { 
                    sheet.Cells[j+2, i+2] = dsBPDetail.Tables[0].Rows[i][j].ToString();
                }                
            }
        
            excelWorkbook.Save();
        }

    /*    public void updateComponentDetail_new(string form_factor)
        {
            string strQuery = " exec sp_app_NPIMRP_compDetail '" + txtBuildID.Text + "','" + form_factor + "'";
            DataSet dsBPDetail = getDataSet(strQuery);
            int col = 0;
            concat_build_id = "";
            if (dsBPDetail.Tables.Count > 2)
            {
                try
                {
                    for (int i = 0; i < dsBPDetail.Tables[2].Rows.Count; i++)
                    {
                        //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());
                        //   concat_build_id = concat_build_id + "," + dsBPDetail.Tables[2].Rows[i]["BP_BuildID"].ToString();
                        for (int j = 2; j < dsBPDetail.Tables[2].Columns.Count; j++)
                        {
                            if (j > 6)
                                col = j + 1;
                            else
                                col = j - 1;
                            if (i == 0)
                                sheet.Cells[i + 17, col] = dsBPDetail.Tables[2].Columns[j].ColumnName.ToString();
                            sheet.Cells[i + 18, col] = dsBPDetail.Tables[2].Rows[i][j].ToString();
                        }
                    }
                }
                catch (Exception e)
                {
                    //MessageBox.Show(e.Message.ToString() );
                }
            }
            excelWorkbook.Save();
        }*/

        public void updateComponentDetail_new(DataTable dtBPDetail)
        {
         //   string strQuery = " exec sp_app_NPIMRP_compDetail '" + txtBuildID.Text + "','" + form_factor + "'";
        //    DataSet dsBPDetail = getDataSet(strQuery);
            int col = 0;
            concat_build_id = "";
           
                try
                {
                    for (int i = 0; i < dtBPDetail.Rows.Count; i++)
                    {
                        //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());
                        //   concat_build_id = concat_build_id + "," + dsBPDetail.Tables[2].Rows[i]["BP_BuildID"].ToString();
                        for (int j = 2; j < dtBPDetail.Columns.Count; j++)
                        {
                            if (j > 6)
                                col = j + 1;
                            else
                                col = j - 1;
                            if (i == 0)
                                sheet.Cells[i + 17, col] = dtBPDetail.Columns[j].ColumnName.ToString();
                            sheet.Cells[i + 18, col] = dtBPDetail.Rows[i][j].ToString();
                        }
                    }
                }
                catch (Exception e)
                {
                    //MessageBox.Show(e.Message.ToString() );
                }
           
            excelWorkbook.Save();
        }
        public void updateCompPOInformation()
        {

          string  filterBy = "BUILD ID";
           string strQuery = "exec sp_NPIMRP_PO_Final '" + build_id + "'";
          // string strQuery = "select po.* from m_npi_mrp_po_all po inner join  (	select distinct comp_pn as comp_pn ,site as site from  dbo.m_npi_mrp_comp_details where build_id in ('3252_16015','3252_16016','3252_16017','3252_16018','3252_16019','3252_16020','3252_16021') )  comp on po.site=comp.site and po.comp_pn=comp.comp_pn ";
         
            bool InRange = false;
            DataSet dsCompDetail = getDataSet(strQuery);

            for (int col = 0; col < dsCompDetail.Tables[0].Columns.Count-1; col++)
            {
                sheet.Cells[1, col+1] = dsCompDetail.Tables[0].Columns[col].ColumnName.ToString();
            }
            int row = 1;
            for (int i = 0; i < dsCompDetail.Tables[0].Rows.Count ; i++)
            {
                row = row + 1;
                 for (int j = 0; j < dsCompDetail.Tables[0].Columns.Count-1; j++)
                    // excelApp.get_Range("A1:A360,B1:E1", Type.Missing).Merge(Type.Missing);
                    sheet.Cells[row,j+1] = dsCompDetail.Tables[0].Rows[i][j].ToString();
            }         
           
            excelWorkbook.Save();
        }


        public void dumpPOInExcel(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
          //    Microsoft.Office.Interop.Excel.Workbook wk =new Microsoft.Office.Interop.Excel.Workbook();
            string constr = "Provider=SQLOLEDB;Data Source=MVerp;Initial Catalog= MIMDist;Trusted_Connection=yes;";
  //     OleDbConnection cn  = new OleDbConnection();
        
   //    Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application();

   //    filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + filename;

   //     String workbookPath   =filename;
       

  //      wk = excelapp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        //wk.SaveAs(filename,  System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value)
   //     Microsoft.Office.Interop.Excel.Worksheet ws  = wk.Sheets[3];

            string strQuery = "exec sp_NPIMRP_PO_Final '" + build_id + "'";

            try
            {
                Microsoft.Office.Interop.Excel.QueryTable oQryTable = ws.QueryTables.Add("OLEDB;" + constr, ws.Range["A1"], strQuery);
                oQryTable.RefreshStyle = Microsoft.Office.Interop.Excel.XlCellInsertionMode.xlOverwriteCells; //' xlInsertEntireRows = 2
                oQryTable.Refresh(false);
                /*  '   excelapp.Visible = True       'set to 'true' when debbugging, Exec is visible
                  '   excelapp.DisplayAlerts = True
                  ' excelapp.SaveWorkspace(workbookPath)
                  'excelapp.MacroOptions()*/
                excelWorkbook.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        finally
            {
           // wk.Close(true, workbookPath, Type.Missing);
           // System.Runtime.InteropServices.Marshal.ReleaseComObject(wk);
          //  System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp);
            }

           

        }

        public void updateAssemblyDetail(string form_factor)
        {
            String strQuery = "select buildID,AsmPartNum, AsmDate,AsmYearWeek,AsmQty from NPIMRP_AsmDetails where plangroup='" + form_factor + "'";
            if (filterBuildID != "")
            {
                strQuery = strQuery + "  and  " + filterBuildID;
            }
            DataSet dsAsmDetail = getDataSet(strQuery + "  Order by buildID  ");
            int col = 0;
                 
            for (int i = 0; i < dsAsmDetail.Tables[0].Rows.Count; i++)
            {
                //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());
                for (int j = 0; j < dsAsmDetail.Tables[0].Columns.Count; j++)
                {
                    sheet.Cells[j + 10, i + 2] = dsAsmDetail.Tables[0].Rows[i][j].ToString();
                }
            }
            if (dsAsmDetail.Tables[0].Rows.Count == 0)
            {
                sheet.Delete();
            }
            else
            {
                sheetnum = sheetnum + 1;              
            }

            excelWorkbook.Save();
        }

      
        //public void getWeekRange(string form_factor )
        //{
        //    string qry = "select cast(dbo.udf_GetISOWeekNumberFromDate(dateadd(wk,-2,build_date)) as varchar)+'-'+cast(year(dateadd(wk,-2,build_date)) as varchar),cast(dbo.udf_GetISOWeekNumberFromDate(dateadd(wk,-1,build_date)) as varchar)+'-'+cast(year(dateadd(wk,-1,build_date)) as varchar),build_week_year,cast(dbo.udf_GetISOWeekNumberFromDate(dateadd(wk,1,build_date)) as varchar)+'-'+cast(year(dateadd(wk,1,build_date)) as varchar), cast(dbo.udf_GetISOWeekNumberFromDate(dateadd(wk,2,build_date)) as varchar)+'-'+cast(year(dateadd(wk,-2,build_date)) as varchar) from dbo.m_npi_mrp_asm_detials   where  " +filterBuildID + " and form_factor='"+form_factor+"'";

        //    DataSet dsWeekRange = getDataSet(qry);
        //    //string form_factor =dsWeekRange.Tables[0].Rows[0]["form_factor"].ToString();
                  
        //}

      
        
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
        string filefor = "";
        DriveService service;

   private void btnUpload_Click(object sender, EventArgs e)
        {

   /*         Cursor.Current = Cursors.WaitCursor;
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

            string promptValue = Prompt.ShowDialog("Please enter the collection name ", "Collection Name");
            string Parentid = "";
            FilesResource.ListRequest lstreq = service.Files.List();
            do
            {
                try
                {
                    FileList fileList = lstreq.Execute();
                    foreach (Google.Apis.Drive.v2.Data.File fileitem in fileList.Items)
                    {
                        if (fileitem.Title == promptValue)
                        {
                            Parentid = fileitem.Id;
                            break;
                        }

                    }
                    lstreq.PageToken = fileList.NextPageToken;
                }
                catch (Exception exp)
                {
                    Console.WriteLine("An error occurred: " + exp.Message);
                    lstreq.PageToken = null;
                }
            } while (!String.IsNullOrEmpty(lstreq.PageToken));*/
   
    /*            File body = new File();               
                body.Title = System.IO.Path.GetFileName(filename);//dialog.FileName);
                body.Description = "document";
                body.MimeType = "application/xlsx";                   
                body.Parents = new List<ParentReference>() { new ParentReference() { Id = Parentid } };

                filename = filename.Replace("file:\\", "");
                //MessageBox.Show("filename=" + filename);
                System.IO.Stream fileStream = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                byte[] byteArray = new byte[fileStream.Length];
                fileStream.Read(byteArray, 0, (int)fileStream.Length);
                System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);               
                FilesResource.InsertMediaUpload request = dr.Files.Insert(body, stream, "application/x-vnd.oasis.opendocument.spreadsheet");
                request.Convert = true;
                request.Upload();
                MessageBox.Show("File has been uploaded in shared doc"); */                     
                Cursor.Current = Cursors.Default;
            }
 

        public string getCollectionName()
        {
            string coll_name = "";
           // string sql = "select * from dbo.vc_vend_master where vend_id=" + cmbVendorFileCreate.SelectedValue;
            //DataSet ds = getdataSet(sql);
            //if (ds != null)
            //    if (ds.Tables.Count > 0)
            //        if (ds.Tables[0].Rows.Count > 0)
            //            coll_name = ds.Tables[0].Rows[0]["shared_doc_name"].ToString();
            return "NPI MRP";
        }
     

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                txtBuildID.Enabled = true;
                txtProjectName.Enabled = false;
            }
            else
            {
                txtBuildID.Enabled = false;
                txtProjectName.Enabled = true;
            }
        }

        private void NPI_MRP_Load(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
