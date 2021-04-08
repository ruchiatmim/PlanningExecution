using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

using System.Configuration;

namespace Version3
{
    public partial class Vendor_Protal : Form
    {
        public Vendor_Protal()
        {
            InitializeComponent();
        }
        string filename = "";
        private void btnExport_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            filename =  this.vendor.Text + "_" + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Year.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + ".xlsx";
            CreateExcel();
            cellUpdate();
            Cursor.Current = Cursors.Default;
        }

        public void CreateExcel()
        {
            //  excelWorkbook.SaveAs("E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\"+filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            //string workbookPath = "E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\VendorSheet.xls";
            filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + filename;

            if (System.IO.File.Exists(filename))
                System.IO.File.Delete(filename);

            string workbookPath = "\\\\mvfile\\MV REPORTS\\Master Data & Forms\\PlanningExecutionForms\\Templates\\Vendor Portal Template.xlsx";
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

        public DataSet getDataSet(string sqlquery, string constr)
        {
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlquery, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;
        }

        public void bindVendor()
        {
            string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
            string strqry = " SELECT distinct vendor_code FROM [MIMDIST].[dbo].[apmaster] where status_type=5 and [vend_class_code]='VENDOR'      SELECT distinct part_no FROM [MIMDIST].[dbo].[inv_master] WHERE     (void = 'N') AND (buyer = 'SOI') AND (obsolete = 0)";
            DataSet ds = getDataSet(strqry, constr);
            vendor.DataSource = ds.Tables[0];
            vendor.ValueMember = "vendor_code";
            vendor.DisplayMember = "vendor_code";           
        }

        public int sheetnum = 0;
        Microsoft.Office.Interop.Excel.Application excelApp1;
        public void cellUpdate()
        {

            string part_num = "";            

                System.Diagnostics.Process p = new System.Diagnostics.Process();
                try
                {
                    excelApp1 = new Microsoft.Office.Interop.Excel.Application();
                    //string workbookPath = "E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\VendorSheet.xls"; Microsoft.Office.Interop.Excel.XlPlatform.xlWindows    
                    string workbookPath = filename;

                    Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp1.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    sheetnum = 0;
                    Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbook.Worksheets;
                    sheetnum = sheetnum + 1;

                    excelApp1.DisplayAlerts = false;
                    Microsoft.Office.Interop.Excel.Worksheet sheet = excelSheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                    String strusername = Environment.UserName;

                    String strQuery = "SELECT po_num,po_line,po_rev,ship_to,part_num,part_desc,mfg_part_num,commit_flag,request_date,ord_qty,due_qty,release_id,release_status,commit_qty,ship_confirm_flag,'' as shipped_qty,carrier,ship_method,tracking_num  FROM [MIMDIST].[dbo].[m_app_vendor_portal] ";

                    string vendor_no = this.vendor.Text;
                    string qtr = this.ComboQtr.Text;
                    if (vendor_no == "")
                        MessageBox.Show("Please select the vendor");
                    else
                    {
                        strQuery = strQuery + "  where   [vendor_code]='" + vendor_no + "' ";// and mim_pn in ('" + part_num.Substring(1, part_num.Length - 1).Replace(",", "','").ToString() + "')";
                    }
                  //  if (qtr != "")
                    //    strQuery = strQuery + " and QUARTER='" + qtr + "'    order by [CUSTOMER PART NUM] ";


                   // sheet.Cells[3, 5] = qtr;
                   // strQuery = strQuery + " SELECT [first_name] + ' '+[last_name] as user_name ,[phone]  ,[email]   FROM [MIMDIST].[dbo].[m_mim_user_table] where user_name='" + strusername + "'";
                    string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
                    DataSet dsRFQ = getDataSet(strQuery, constr);
                    int col = 0;
                    for (int i = 0; i < dsRFQ.Tables[0].Rows.Count; i++)
                    {
                        int row = i + 2;
                        for (int j = 0; j < 18; j++)
                        {
                            col = j + 1;

                            // excelApp.get_Range("A1:A360,B1:E1", Type.Missing).Merge(Type.Missing);
                            sheet.Cells[row, col] = dsRFQ.Tables[0].Rows[i][j].ToString();
                        }
                    }

                /*    if (dsRFQ.Tables[1].Rows.Count > 0)
                    {
                        sheet.Cells[1, 10] = dsRFQ.Tables[1].Rows[0][0].ToString();
                        sheet.Cells[2, 10] = dsRFQ.Tables[1].Rows[0][1].ToString();
                        sheet.Cells[3, 10] = dsRFQ.Tables[1].Rows[0][2].ToString();
                    }*/
                    excelWorkbook.Save();
                    excelWorkbook.Close(true, false, Type.Missing);

                    p.StartInfo.FileName = filename;
                    //p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
                    p.Start();

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

        private void Vendor_Protal_Load(object sender, EventArgs e)
        {
            bindVendor();
        }
    }
}
