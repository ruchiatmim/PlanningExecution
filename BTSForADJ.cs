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
using Infragistics.Win.UltraWinGrid;


namespace Version3
{
    public partial class BTSForADJ: Form
    {
        public BTSForADJ()
        {
            InitializeComponent();
        }
        string constr = ConfigurationManager.ConnectionStrings["MVIntConnectionString"].ConnectionString;
        private void BTSForMT_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'bTS.BTS_GoogleIntInvTranWB' table. You can move, or remove it, as needed.
           // this.bTS_GoogleIntInvTranTableAdapter.Fill(this.bTS.BTS_GoogleIntInvTranWB);
            this.txtStartdate.Value = DateTime.Now.AddDays(-3);
            this.txtEndDate.Value = DateTime.Now;
          //  loadStatus();
            LoadDropDown();
            loadgrid();

        }

     

        DataTable dtStatus;
        DataTable dtTranType;
        DataTable dtApp;
        DataTable dtRegion;
        DataTable dtTranSource;
        DataTable dtFolder;
        DataTable dtSite;
        private void LoadDropDown()
        {
            dtStatus = new DataTable();
            dtStatus.Columns.Add("Status");
            dtStatus.Columns.Add("Value");
            dtStatus.Rows.Add("AUDIT", "A");
            dtStatus.Rows.Add("HOLD", "H");
            dtStatus.Rows.Add("NEW", "N");
            dtStatus.Rows.Add("PROCCESSED", "Y");
            dtStatus.Rows.Add("RESOLVED", "X");
            dtStatus.Rows.Add("VOID", "V");
           

            ddStatus.DataSource = dtStatus;
            ddStatus.DisplayMember = "Status";
            ddStatus.ValueMember = "Value";

            dtStatus.Rows.Add("ALL", "");
            cmbStatus.DataSource = dtStatus;
            cmbStatus.DisplayMember = "Status";
            cmbStatus.ValueMember = "Value";
            cmbStatus.SelectedText = "ALL";

            dtTranType = new DataTable();
            dtTranType.Columns.Add("Tran Type");

          //  dtTranType.Rows.Add("ASSEMBLY COMPLETION");
          //  dtTranType.Rows.Add("COMPONENT ISSUE");
          //  dtTranType.Rows.Add("COMPONENT RETURN");
         //   dtTranType.Rows.Add("DECOM2");
         //   dtTranType.Rows.Add("DECOM2 ISSUE");
         //   dtTranType.Rows.Add("DECOM2 RETURN");
            dtTranType.Rows.Add("MATERIAL RECEIPT");
            dtTranType.Rows.Add("MATERIAL TRANSFER");

            this.ddTranType.DataSource = dtTranType;
            ddTranType.DisplayMember = "Tran Type";
            ddTranType.ValueMember = "Tran Type";
            
        

            DataTable dtSiteFilter = new DataTable();

            dtSiteFilter.Columns.Add("Site");
            dtSiteFilter.Rows.Add("ALL");
            dtSiteFilter.Rows.Add("FB");
            dtSiteFilter.Rows.Add("NF");



            cmbPlant.DataSource = dtSiteFilter;
            cmbPlant.DisplayMember = "Site";
            cmbPlant.ValueMember = "Site";
            cmbPlant.SelectedText = "ALL";

            dtRegion = new DataTable();
            dtRegion.Columns.Add("Region");
            dtRegion.Rows.Add("112");
            dtRegion.Rows.Add("113");
            this.ddRegion.DataSource = dtRegion;

            dtTranSource = new DataTable();
            dtTranSource.Columns.Add("TranSrc");
            dtTranSource.Rows.Add("MT");
            dtTranSource.Rows.Add("TWO");
            this.ddSource.DataSource = dtTranSource;

            dtFolder = new DataTable();
            dtFolder.Columns.Add("Folder");
            dtFolder.Rows.Add("MIMIRE");
            dtFolder.Rows.Add("MIMUS");
            this.ddFolder.DataSource = dtFolder;

            dtSite = new DataTable();
            dtSite.Columns.Add("loc");
            dtSite.Rows.Add("AT");
            dtSite.Rows.Add("NL");
    
            this.ddLocation.DataSource = dtSite;

            DataTable dtEDIRef = new DataTable();
            dtEDIRef.Columns.Add("EDI Reference");
            dtEDIRef.Rows.Add("947");

            this.ddEDIRef.DataSource = dtEDIRef;   
    
            DataTable dtFromToLoc = new DataTable();
            dtFromToLoc.Columns.Add("Location");
            dtFromToLoc.Rows.Add("CMMIMG1E");
            dtFromToLoc.Rows.Add("CMMIMG1P");
            dtFromToLoc.Rows.Add("CMMIMG1PC");
            dtFromToLoc.Rows.Add("CMMIMN2E");
            dtFromToLoc.Rows.Add("CMMIMN2P");           
            dtFromToLoc.Rows.Add("CMMIMN2PC");
            dtFromToLoc.Rows.Add("CMMIMN2PK");         
            dtFromToLoc.Rows.Add("CYCLECOUNT");
            dtFromToLoc.Rows.Add("SN_CHANGE");
            
            
            
    
            this.ddFromToLoc.DataSource = dtFromToLoc;
          
   
        }

        public void getSerialNum(string Ident)
        {
            
            //WebDataGridSerial.DataSource = null;
            string sql = "  SELECT [AutoId],Status,[SerialNum],[ParentSerialNum] ,[AssetTagNum],Reference  FROM [dbo].[BTS_GoogleIntInvTranSerial] where uniqueIdent='" + Ident + "'";
            SqlConnection cn = new SqlConnection(constr);
            SqlCommand cmd = new SqlCommand(sql, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet dsSerial = new DataSet();
            da.Fill(dsSerial);
            ugBTSSerial.DataSource = dsSerial.Tables[0];                     

        }

        private void ugBTS_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {

            //  e.Layout.Bands[0].Columns["Status"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
            e.Layout.Bands[0].Columns["EDIStatus"].ValueList = this.ddStatus;
            //     e.Layout.Bands[0].Columns["TranType"].ValueList = this.ddTranType;
            //   e.Layout.Bands[0].Columns["Region"].ValueList = this.ddRegion;
            //    e.Layout.Bands[0].Columns["Folder"].ValueList = this.ddFolder;

            //    e.Layout.Bands[0].Columns["TranSource"].ValueList = this.ddSource;
            //   e.Layout.Bands[0].Columns["app"].ValueList = this.ddApp;
            //    e.Layout.Bands[0].Columns["location"].ValueList = this.ddLocation;

        }
        string uniqueIdent = "";
        private void ugBTS_AfterSelectChange(object sender, Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs e)
        {
            if (this.ugBTS.Selected.Rows.Count > 0)
            {
                uniqueIdent = ugBTS.Selected.Rows[0].Cells["uniqueIdent"].Value.ToString().Trim();
            }
            else
            {
                uniqueIdent = "";
            }
            txtUniqueIdentifier.Text = uniqueIdent;
            getSerialNum(uniqueIdent);
        
        }
        string sql;
        public void loadgrid()
        {
            Cursor.Current = Cursors.WaitCursor;
          //  sql = "select  [auto_id],[status],Replace(replace(Replace(Replace([filename],'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMUS\\outbound\\TEMP\\',''),'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMUS\\outbound\\INVTRAN\\',''),'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMIRE\\outbound\\TEMP\\',''),'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMIRE\\outbound\\INVTRAN\\','') as filename,[fileDate],[ProcessDate],[dekit_summary_id] ,[tran_id],[uniqueIdent] ,[from_loc] ,[to_loc] ,[from_proj],[to_proj],[tran_date],[part_no],[tran_type],[tran_source],[tran_source_num] ,[tran_source_line],[serial_num] ,[qty] ,[reference]  ,[location]  from BTS_GoogleIntInvTran where  uniqueIdent is not null  ";  //,replace(filename,'\\\\mimsftp\\e$\\sftp\\msgGoogle\\PLATFORM\\PROD\\','') as file1

            sql = "SELECT [AutoId]         ,[TranId]   ,UniqueIdent        ,[FromLoc]      ,[ToLoc]      ,[FromProj]      ,[ToProj]      ,[TranDate]      ,[PartNo]  ,    [TranType]        ,[TranSource]      ,[TranSourceNum]      ,[TranSourceLine]         ,[Qty]      ,[Reference]                  ,[Plant]      ,EDIStatus   ,[UTCdate]      ,[PoNum]      ,[PoLine],ICNSentDate,[EDIICN#]  as ICN,EDIReference   FROM  [dbo].[BTS_GoogleIntInvTranForADJ]    where  uniqueIdent is not null     ";
         

            if (txtFromLoc.Text.Trim() != "")
                sql = sql + " and   FromLoc like '" + txtFromLoc.Text.Trim() + "%'";
            if (txtToLoc.Text.Trim() != "")
                sql = sql + " and   ToLoc like '" + txtToLoc.Text.Trim() + "%'";

            if (cmbStatus.Text != "ALL")
            {
                sql = sql + " and   EDIStatus='" + cmbStatus.Value + "'";
            }

            if (cmbPlant.Text != "ALL")
            {
                sql = sql + " and   Plant='" + cmbPlant.Value + "'";
            }

            if (this.txtPartNum.Text.Trim() != "")
                sql = sql + " and   [PartNo] like '" + txtPartNum.Text.Trim() + "%'";

            if (this.txtTranSourceNum.Text.Trim() != "")
                sql = sql + " and   [TranSourceNum] like '" + txtTranSourceNum.Text.Trim() + "%'";

            if (this.txtStartdate.Value != "" && this.txtEndDate.Value != "")
            {
                DateTime st_date = Convert.ToDateTime(Convert.ToDateTime(this.txtStartdate.Value).ToShortDateString());
                DateTime end_date = Convert.ToDateTime(Convert.ToDateTime(this.txtEndDate.Value).AddDays(1).ToShortDateString());
                sql = sql + " and   [TranDate]>='" + st_date + "'  and [TranDate]<='" + end_date + "'  ";
            }
            else
            {
                MessageBox.Show("Please select start tran date and end tran date");
            }

            sql = sql + "  order by [AutoId] desc";
            //   SqlDataSource1.SelectCommand = sql;

            SqlConnection cn = new SqlConnection(constr);
            SqlCommand cmd = new SqlCommand(sql, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            ugBTS.DataSource = ds.Tables[0];
            ugBTS.DisplayLayout.Override.CellClickAction = CellClickAction.EditAndSelectText;

            lblnumRec.Text = "Number of Records :" + ds.Tables[0].Rows.Count.ToString();
            //this.ddSource.DataSource = sourceTable;
            //this.ddRegion.DataSource = RegionTable;
            //this.ddFolder.DataSource = folderTable;
            //this.ddLocation.DataSource = LocTable;
            Cursor.Current = Cursors.Default;
        }

        private void ultraButton1_Click(object sender, EventArgs e)
        {
            loadgrid();
            //Cursor.Current = Cursors.WaitCursor;
            //string sql = "select  [auto_id],[status],Replace(replace(Replace(Replace([filename],'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMUS\\outbound\\TEMP\\',''),'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMUS\\outbound\\INVTRAN\\',''),'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMIRE\\outbound\\TEMP\\',''),'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMIRE\\outbound\\INVTRAN\\','') as filename,[fileDate],[ProcessDate],[dekit_summary_id] ,[tran_id],[uniqueIdent] ,[from_loc] ,[to_loc] ,[from_proj],[to_proj],[tran_date],[part_no],[tran_type],[tran_source],[tran_source_num] ,[tran_source_line],[serial_num] ,[qty] ,[reference]  ,[location]  from BTS_GoogleIntInvTranWB where  uniqueIdent is not null  ";  //,replace(filename,'\\\\mimsftp\\e$\\sftp\\msgGoogle\\PLATFORM\\PROD\\','') as file1
            //if (cmbStatus.Text != "ALL")
            //{
            //    sql = sql + " and   status='" + cmbStatus.Value + "'";
            //}
       
            //if (txtTranID.Text.Trim() != "")
            //    sql = sql + " and   tran_id like '" + txtTranID.Text.Trim() + "%'";

            //if (this.txtUniqueIdent.Text.Trim() != "")
            //    sql = sql + " and   uniqueIdent like '" + txtUniqueIdent.Text.Trim() + "%'";

            ////if (this.txtPartNum.Text.Trim() != "")
            ////    sql = sql + " and   part_no like '" + txtPartNum.Text.Trim() + "%'";

            ////if (this.txtTranSourceNum.Text.Trim() != "")
            ////    sql = sql + " and   tran_source_num like '" + txtTranSourceNum.Text.Trim() + "%'";         

            //if (this.txtStartdate.Value != "" && this.txtEndDate.Value != "")
            //{
            //    DateTime st_date =Convert.ToDateTime( Convert.ToDateTime(this.txtStartdate.Value).ToShortDateString());
            //    DateTime end_date = Convert.ToDateTime( Convert.ToDateTime(this.txtEndDate.Value).ToShortDateString());
            //    sql = sql + " and   tran_date>='" + st_date + "'  and tran_date<='" + end_date + "'  ";
            //}
            //else
            //{
            //    MessageBox.Show("Please select start tran date and end tran date"); 
            //}

            //sql = sql + "  order by auto_id desc";
            ////   SqlDataSource1.SelectCommand = sql;

            //SqlConnection cn = new SqlConnection(constr);
            //SqlCommand cmd = new SqlCommand(sql, cn);
            //SqlDataAdapter da = new SqlDataAdapter(cmd);
            //DataSet ds = new DataSet();
            //da.Fill(ds);
            //ugBTS.DataSource = ds.Tables[0];
            //ugBTS.DisplayLayout.Override.CellClickAction = CellClickAction.EditAndSelectText;

            //lblnumRec.Text = "Number of Records :" + ds.Tables[0].Rows.Count.ToString();
            ////this.ddSource.DataSource = sourceTable;
            ////this.ddRegion.DataSource = RegionTable;
            ////this.ddFolder.DataSource = folderTable;
            ////this.ddLocation.DataSource = LocTable;
            //Cursor.Current = Cursors.Default;
        }


        private DataTable distinctTable(DataTable dt, string ColumnName)           
        {

            DataView view = new DataView(dt);
            DataTable distinctValues = new DataTable();
            distinctValues = view.ToTable(true, ColumnName);
            return distinctValues;

        }

        private void ultraButton2_Click(object sender, EventArgs e)
        {
            
        }

        private void ultraButton3_Click(object sender, EventArgs e)
        {
           
        }

        private void ugBTSSerial_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {

        }

        private void ubDelete_Click(object sender, EventArgs e)
        {
           
        }

        private void ubAddNew_Click(object sender, EventArgs e)
        {
           // this.ugBTSSerial.DisplayLayout.Bands[0].AddNew();
        }

        private void ubDelete_Click_1(object sender, EventArgs e)
        {
            //int serial_auto_id = 0;
            //if (this.ugBTSSerial.Selected.Rows.Count > 0)
            //{
            //    if (this.ugBTSSerial.Selected.Rows[0].Cells[0].Text.ToString().Trim() != "")
            //    {
            //        serial_auto_id = Convert.ToInt32(this.ugBTSSerial.Selected.Rows[0].Cells[0].Text.ToString().Trim());
            //        string deleteQry = "delete from BTS_GoogleIntInvTranSerial where auto_id=" + serial_auto_id + "  select @@error ";
            //        SqlConnection cn = new SqlConnection(constr);
            //        cn.Open();
            //        SqlCommand cmd = new SqlCommand(deleteQry, cn);
            //        SqlDataAdapter da = new SqlDataAdapter(cmd);
            //        DataSet ds = new DataSet();
            //        da.Fill(ds);
            //        cn.Close();
            //        if (ds != null)
            //            if (ds.Tables.Count > 0)
            //                if (ds.Tables[0].Rows.Count > 0)
            //                {
            //                    if (ds.Tables[0].Rows[0][0].ToString() == "0")
            //                        MessageBox.Show("Serial Num has been deleted Successfully");
            //                }
            //        }
           // }
        }

  
        private void ugBTS_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
        {
            //int auto_id = 0;
            //string reference = "";
            //string status = "";
            //if (e.Cell.Column.Index == 1 || e.Cell.Column.Index == 20)
            //{
              
            //  //  if (this.ugBTS.Selected.Rows.Count > 0)
            //  //  {
            //    auto_id = Convert.ToInt32(e.Cell.Row.Cells["auto_id"].Value.ToString().Trim());
            //    status = e.Cell.Row.Cells["status"].Value.ToString().Trim();
            //    reference = e.Cell.Row.Cells["reference"].Value.ToString().Trim();
            //    string qry = "update dbo.BTS_GoogleIntInvTranWB set status='" + status + "', reference='" + reference + "'   where auto_id=" + auto_id;
            //    SqlConnection cn = new SqlConnection(constr);
            //    cn.Open();
            //    SqlCommand cmd = new SqlCommand(qry, cn);
            //    SqlDataAdapter da = new SqlDataAdapter(cmd);
            //    DataSet ds = new DataSet();
            //    da.Fill(ds);
            //    cn.Close();
            //   // }
            //}                    
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
           

        }

        private void lnkInsertblank_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DataTable dt = ugBTS.DataSource as DataTable;
            DataRow dr = dt.NewRow();
         //   for (int i = 1; i < ugBTS.Selected.Rows[0].Cells.Count; i++)
           //     dr[i] = ugBTS.Selected.Rows[0].Cells[i].Value;
            dr["autoid"] = 0;
            dt.Rows.Add(dr);
            ugBTS.DataSource = dt;
            ugBTS.Rows[ugBTS.Rows.Count - 1].Activate();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DataTable dt = ugBTS.DataSource as DataTable;
            DataRow dr = dt.NewRow();
            if (ugBTS.Selected.Rows.Count > 0)
            {
                dr["autoid"] = 0;
             //   dr["status"] = "N";
                for (int i = 1; i < ugBTS.Selected.Rows[0].Cells.Count - 4; i++)

                    if (ugBTS.Selected.Rows[0].Cells[i].Column.Key.ToUpper() == "EDISTATUS")
                        dr[i] = "N";
                    else if (ugBTS.Selected.Rows[0].Cells[i].Column.Key.ToUpper() == "UNIQUEIDENT")
                        dr[i] = "";
                    else 
                        dr[i] = ugBTS.Selected.Rows[0].Cells[i].Value;                   
                 
                dt.Rows.Add(dr);
                ugBTS.DataSource = dt;
                ugBTS.Rows[ugBTS.Rows.Count - 1].Activate();
            }
            else
                MessageBox.Show("Please select row");
        }

        private void ultraButton2_Click_1(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
          
            if (System.IO.File.Exists(path+"\\GridData.xlsx"))
            {
                try
                {
                    System.IO.File.Delete(path + "\\GridData.xlsx");
                }

                catch (Exception exp)
                {
                    string Msg = exp.Message.ToString();
                    MessageBox.Show("Please close file GridData.xlsx then Export", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    goto endprg;
                }

              
            }
            if (this.ugBTS.Rows.Count > 0)
            {
                // this.ultraGridExcelExporter1.Export(this.ugBTS, path + "\\GridData.xlsx");
                DataTable dtGrid = ugBTS.DataSource as DataTable;
              
                exportToExcel1(path+"\\GridData.xlsx");
            }
            else
                MessageBox.Show("There is no row in the grid.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
          
           endprg:
            Cursor.Current = Cursors.Default;
        }

        private void exportToExcel1(string filename)
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            Microsoft.Office.Interop.Excel.Application excelApp1 = new Microsoft.Office.Interop.Excel.Application();
            try
            {
            

            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp1.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            Microsoft.Office.Interop.Excel.Worksheet sheet = excelSheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
            string cnt = "data source=mvint;initial catalog=proddist;user id=sa;password=mimi~100;pooling=true;max pool size=100;min pool size=1;";
            Microsoft.Office.Interop.Excel.QueryTable oQryTable = sheet.QueryTables.Add("OLEDB;Provider=sqloledb;" + cnt, sheet.Range["A1"], sql);
            oQryTable.RefreshStyle = Microsoft.Office.Interop.Excel.XlCellInsertionMode.xlInsertEntireRows; // 2; //' xlInsertEntireRows = 2
            oQryTable.Refresh(false);
            excelApp1.DisplayAlerts = false;
            excelWorkbook.Save();
            excelWorkbook.SaveAs(filename);
       
            p.StartInfo.FileName = filename;
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


        public void exportToExcel(DataTable dt,string filename)
        {
          
            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp1 = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp1.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbook.Worksheets;
                Microsoft.Office.Interop.Excel.Worksheet sheet = excelSheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                int col = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //copyPasteRangeExcel(("C" + j).ToString(), ("AL" + j).ToString());
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        sheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                    }
                }
                excelWorkbook.Save();

                excelWorkbook.SaveAs(filename);
            }
            catch (Exception e)
            {

            }
           
            
        }
        private void ddStatus_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {

            e.Layout.ScrollStyle = ScrollStyle.Immediate;

            // Change the order in which columns get displayed in the UltraDropDown.        
            // Hide columns you don't want shown.
            e.Layout.Bands[0].Columns[1].Hidden = true;

            // Sort the items in the drop down by ProductName column.
            //      e.Layout.Bands[0].SortedColumns.Clear();
            //  e.Layout.Bands[0].SortedColumns.Add("Status", false);

            // Set the border style of the drop down.
            e.Layout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;	
		

        }

        private void cmbStatus_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.ScrollStyle = ScrollStyle.Immediate;

            // Change the order in which columns get displayed in the UltraDropDown.        
            // Hide columns you don't want shown.
            e.Layout.Bands[0].Columns[1].Hidden = true;

            // Sort the items in the drop down by ProductName column.
            //      e.Layout.Bands[0].SortedColumns.Clear();
            //  e.Layout.Bands[0].SortedColumns.Add("Status", false);

            // Set the border style of the drop down.
            e.Layout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;	
		

        }

        private void ugBTS_BeforeCellListDropDown(object sender, CancelableCellEventArgs e)
        {
            UltraDropDown udd = e.Cell.ValueListResolved as UltraDropDown;
            if (udd != null)
                udd.DropDownWidth = e.Cell.Column.CellSizeResolved.Width;
        }

        private void ugBTS_ClickCell(object sender, ClickCellEventArgs e)
        {
                int var_auto_id;
                string var_status;
                int var_tran_id;

                string var_from_loc;
                string var_to_loc;
                string var_from_proj;
                string var_to_proj;
                DateTime var_tran_date;
                DateTime var_utc_date;
                string var_part_no;
                string var_tran_type;
                string var_tran_source;
                string var_tran_source_num;
                string var_tran_source_line;
                string var_serial_num;
                int var_qty;
                string var_reference;
                string var_location;
               string  var_ponum;
                int var_poline;
                var_auto_id = 0;
                var_ponum="";
                var_poline = 0;
                string var_EDIStatus;
                string var_EDIreference;
            if (e.Cell.Column.Key == "Save")
            {
                if  (e.Cell.Row.Cells["autoid"].Value !="" && e.Cell.Row.Cells["autoid"].Value!=null)
                var_auto_id = Convert.ToInt32(e.Cell.Row.Cells["autoid"].Value);
                var_status ="N";
                var_tran_id = Convert.ToInt32(e.Cell.Row.Cells["tranid"].Value.ToString());
                var_from_loc = e.Cell.Row.Cells["fromloc"].Value.ToString();
                var_to_loc = e.Cell.Row.Cells["toloc"].Value.ToString();
                var_from_proj = e.Cell.Row.Cells["fromproj"].Value.ToString();
                var_to_proj = e.Cell.Row.Cells["toproj"].Value.ToString();
                var_tran_date = Convert.ToDateTime(e.Cell.Row.Cells["trandate"].Value.ToString());
                var_utc_date = Convert.ToDateTime(e.Cell.Row.Cells["utcdate"].Value.ToString());
                var_part_no = e.Cell.Row.Cells["partno"].Value.ToString();
                var_tran_type = e.Cell.Row.Cells["trantype"].Value.ToString();
                var_tran_source = e.Cell.Row.Cells["transource"].Value.ToString();
                var_tran_source_num = e.Cell.Row.Cells["transourcenum"].Value.ToString();
                var_tran_source_line = e.Cell.Row.Cells["transourceline"].Value.ToString();
                var_serial_num ="";
                var_qty = Convert.ToInt32(e.Cell.Row.Cells["qty"].Value.ToString());
                var_reference = e.Cell.Row.Cells["reference"].Value.ToString();
                var_location = e.Cell.Row.Cells["Plant"].Value.ToString();
                var_EDIreference = e.Cell.Row.Cells["EDIReference"].Value.ToString();
                var_EDIStatus = e.Cell.Row.Cells["EDIStatus"].Value.ToString();
             //    if  (e.Cell.Row.Cells["ponum"].Value !="" && e.Cell.Row.Cells["ponum"].Value!=null)
              //  var_ponum = Convert.ToInt32(e.Cell.Row.Cells["ponum"].Value.ToString());  
             //    if  (e.Cell.Row.Cells["poline"].Value !="" && e.Cell.Row.Cells["poline"].Value!=null)    
            //    var_poline =  Convert.ToInt32(e.Cell.Row.Cells["poline"].Value.ToString());

                string sqlstr = "exec [sp_updateBTSForADJ] '" + var_auto_id.ToString() + "','" + var_status + "','" + var_tran_id + "','" + var_from_loc + "','" + var_to_loc + "','" + var_from_proj + "','" + var_to_proj + "','" + var_tran_date + "','" + var_part_no + "','" + var_tran_type + "','" + var_tran_source + "','" + var_tran_source_num + "','" + var_tran_source_line + "','" + var_serial_num + "','" + var_qty + "','" + var_reference + "','" + var_location + "','0','0','" + var_EDIreference + "','" + var_EDIStatus + "','" + var_utc_date + "'";
                UpdateRow(sqlstr);
            }
        }

        public void UpdateRow(string sql)
        {           
            
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sql, cn);
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                //  tran.Commit();
                MessageBox.Show("Updated successfully");
            }

            catch (Exception exp)
            {
                string error = exp.Message.ToString();
                // tran.Rollback();
                MessageBox.Show("Failed to Update. Please check the data");
            }
            loadgrid();
            cn.Close();
        }

        private void ubAddNew_Click_1(object sender, EventArgs e)
        {
            this.ugBTSSerial.DisplayLayout.Bands[0].AddNew();
        }

        private void ubSave_Click(object sender, EventArgs e)
        {
            string qry = "";
            //MessageBox.Show(vendorGrid.Rows.Count.ToString());
            for (int i = 0; i < ugBTSSerial.Rows.Count; i++)
            {
                int auto_id = 0;
                string serial_num = ugBTSSerial.Rows[i].Cells["serialnum"].Value.ToString().Trim();
                string parent_serial_num = ugBTSSerial.Rows[i].Cells["parentserialnum"].Value.ToString().Trim();
                // string asset_tag_num = ugBTSSerial.Rows[i].Cells["asset_tag_num"].Value.ToString();

                if ((ugBTSSerial.Rows[i].Cells["autoid"].Value != null) && (ugBTSSerial.Rows[i].Cells["autoid"].Value.ToString().Trim() != ""))
                    auto_id = Convert.ToInt32(ugBTSSerial.Rows[i].Cells["autoid"].Value.ToString().Trim());
                string app = "";
                if (serial_num != "")
                    app = "WB";
                else
                    app = "EPI";
                if (auto_id == 0)
                {
                    qry = qry + " insert into BTS_GoogleIntInvTranSerial([Status] ,[TranId],[uniqueIdent] ,[TranDate],[PartNo] ,[SerialNum] ,[ParentSerialNum] ,[Reference]      ,[AssetTagNum]            ,[EPITranType]      ,[Plant]      ,[App]) select  'N', [TranId],[uniqueIdent] ,[TranDate] ,[PartNo] ,'" + serial_num + "','" + parent_serial_num + "',[TranSourceNum],'',[EPITranType]  ,[Plant]  ,[App]   from [dbo].[BTS_GoogleIntInvTranForADJ] where uniqueIdent='" + uniqueIdent + "'";
                }
                else
                {
                    qry = qry + "   update dbo.BTS_GoogleIntInvTranSerial  set  parentserialnum='" + parent_serial_num + "', serialnum='" + serial_num + "' where autoid=" + auto_id.ToString();
                }
            }

            //   string constr = ConfigurationManager.ConnectionStrings["epConnectionStringTest"].ConnectionString;

            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(qry, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            getSerialNum(uniqueIdent);
        }

        private void ubDelete_Click_2(object sender, EventArgs e)
        {
            int serial_auto_id = 0;
            if (this.ugBTSSerial.Selected.Rows.Count > 0)
            {
                if (this.ugBTSSerial.Selected.Rows[0].Cells[0].Text.ToString().Trim() != "")
                {
                    serial_auto_id = Convert.ToInt32(this.ugBTSSerial.Selected.Rows[0].Cells[0].Text.ToString().Trim());
                    string deleteQry = "delete from [dbo].[BTS_GoogleIntInvTranSerial] where autoid=" + serial_auto_id + "  select @@error ";
                    SqlConnection cn = new SqlConnection(constr);
                    cn.Open();
                    SqlCommand cmd = new SqlCommand(deleteQry, cn);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    cn.Close();
                    if (ds != null)
                        if (ds.Tables.Count > 0)
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                if (ds.Tables[0].Rows[0][0].ToString() == "0")
                                    MessageBox.Show("Serial Num has been deleted Successfully");
                            }
                }
            }


        }

        private void btnSaveAll_Click(object sender, EventArgs e)
        {
              int var_auto_id;
                string var_status;
                int var_tran_id;
                string var_from_loc;
                string var_to_loc;
                string var_from_proj;
                string var_to_proj;
                DateTime var_tran_date;
                string var_part_no;
                string var_tran_type;
                string var_tran_source;
                string var_tran_source_num;
                string var_tran_source_line;
                string var_serial_num;
                int var_qty;
                string var_reference;
                string var_EDIreference;
                string var_location;
                int var_ponum;
                int var_poline;
                string var_EDIStatus;
                var_auto_id = 0;
                var_ponum=0;
                var_poline = 0;
             string sqlstr ="";
             foreach (UltraGridRow row in ugBTS.Rows)
             {
                 if (row.Cells["autoid"].Value != "" && row.Cells["autoid"].Value != null)
                     var_auto_id = Convert.ToInt32(row.Cells["autoid"].Value);
                 var_status = "N";
                 var_tran_id = Convert.ToInt32(row.Cells["tranid"].Value.ToString());
                 var_from_loc = row.Cells["fromloc"].Value.ToString();
                 var_to_loc = row.Cells["toloc"].Value.ToString();
                 var_from_proj = row.Cells["fromproj"].Value.ToString();
                 var_to_proj = row.Cells["toproj"].Value.ToString();
                 var_tran_date = Convert.ToDateTime(row.Cells["trandate"].Value.ToString());
                 var_part_no = row.Cells["partno"].Value.ToString();
                 var_tran_type = row.Cells["trantype"].Value.ToString();
                 var_tran_source = row.Cells["transource"].Value.ToString();
                 var_tran_source_num = row.Cells["transourcenum"].Value.ToString();
                 var_tran_source_line = row.Cells["transourceline"].Value.ToString();
                 var_serial_num = "";
                 var_qty = Convert.ToInt32(row.Cells["qty"].Value.ToString());
                 var_reference = row.Cells["reference"].Value.ToString();
                 var_location = row.Cells["Plant"].Value.ToString();
                 var_EDIreference = row.Cells["EDI Reference"].Value.ToString();
                 //    if  (e.Cell.Row.Cells["ponum"].Value !="" && e.Cell.Row.Cells["ponum"].Value!=null)
                 //  var_ponum = Convert.ToInt32(e.Cell.Row.Cells["ponum"].Value.ToString());  
                 //    if  (e.Cell.Row.Cells["poline"].Value !="" && e.Cell.Row.Cells["poline"].Value!=null)    
                 //    var_poline =  Convert.ToInt32(e.Cell.Row.Cells["poline"].Value.ToString());
                 var_EDIStatus = row.Cells["EDIStatus"].Value.ToString();
                 sqlstr = sqlstr + "exec [sp_updateBTSForMT] '" + var_auto_id.ToString() + "','" + var_status + "','" + var_tran_id + "','" + var_from_loc + "','" + var_to_loc + "','" + var_from_proj + "','" + var_to_proj + "','" + var_tran_date + "','" + var_part_no + "','" + var_tran_type + "','" + var_tran_source + "','" + var_tran_source_num + "','" + var_tran_source_line + "','" + var_serial_num + "','" + var_qty + "','" + var_reference + "','" + var_location + "','0','0','" + var_EDIreference + "','" + var_EDIStatus + "'";
             }
             UpdateRow(sqlstr);
        }
        void ob_FormClosed(object sender, FormClosedEventArgs e)
        {
            loadgrid();
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            ManualAdjustment ma = new ManualAdjustment();
            ma.FormClosed += new FormClosedEventHandler(ob_FormClosed);
         //   ma.ButtonClicked += new EventHandler(ob_ButtonClicked);
           
            ma.ShowDialog();
            /*DataTable dt = ugBTS.DataSource as DataTable;
            DataRow dr = dt.NewRow();
         
                dr["autoid"] = 0;
                //   dr["status"] = "N";
                //for (int i = 1; i < ugBTS.Selected.Rows[0].Cells.Count - 4; i++)    
                  
                        dr["EDISTATUS"] = "A";                  
                        dr["UNIQUEIDENT"] = "";
                        dr["Fromloc"] = "GPNChange";
                        dr["partno"] = "";

                        dr["FromProj"] = "MAIN";
                        dr["ToProj"] = "MAIN";

                        dr["TranType"] = "MATERIAL TRANSFER";
                        dr["transource"] = "MT";
                        dr["transourceline"] = "1";
                   
                   
                      //  dr[i] = ugBTS.Selected.Rows[0].Cells[i].Value;

                dt.Rows.Add(dr);
                 dr = dt.NewRow();
          
           
                dr["autoid"] = 0;
                //   dr["status"] = "N";
                //for (int i = 1; i < ugBTS.Selected.Rows[0].Cells.Count - 4; i++)    
                  
                        dr["EDISTATUS"] = "A";                  
                     //   dr["UNIQUEIDENT"] = "";
                        dr["ToLoc"] = "GPNChange";
                        dr["partno"] = "";

                        dr["FromProj"] = "MAIN";
                        dr["ToProj"] = "MAIN";

                        dr["TranType"] = "MATERIAL TRANSFER";
                        dr["transource"] = "MT";
                        dr["transourceline"] = "1";
                   
                   //     dr[i] = ugBTS.Selected.Rows[0].Cells[i].Value;

                dt.Rows.Add(dr);
                ugBTS.DataSource = dt;
              
           
            ugBTS.Rows[ugBTS.Rows.Count - 2].Activate();
             * */
          
        }

        private void button1_Click(object sender, EventArgs e)
        {

          

          //  loadgrid();

           // DataTable dtGPNChange = ugBTS.DataSource as DataTable;
          //  DataRow[] drGPN = dtGPNChange.Select("(FromLoc='GPNChange' or ToLoc='GPNChange') and auto_id=0");
          /*  if (drGPN.Length == 2)
            {

            }*/
           

        }
    }
}
