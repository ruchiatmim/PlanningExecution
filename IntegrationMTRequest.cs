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
    public partial class IntegrationMTRequest : Form
    {
        public IntegrationMTRequest()
        {
            InitializeComponent();
        }
        string constr = ConfigurationManager.ConnectionStrings["MVIntConnectionString"].ConnectionString;
        private void IntegrationMTRequest_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'bTS.BTS_GoogleIntInvTranWB' table. You can move, or remove it, as needed.
           // this.bTS_GoogleIntInvTranTableAdapter.Fill(this.bTS.BTS_GoogleIntInvTranWB);
            this.txtStartdate.Value = DateTime.Now.AddDays(-3);
            this.txtEndDate.Value = DateTime.Now;
          //  loadStatus();
            LoadDropDown();
            loadgrid();

        }

        public void loadStatus()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Status");
            DataRow dr=dt.NewRow();
            dr["Status"]="Y";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["Status"] = "N";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["Status"] = "ALL";
            dt.Rows.Add(dr);
            cmbStatus.DataSource = dt;
            cmbStatus.SelectedText = "ALL";

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

            dtTranType.Rows.Add("ASSEMBLY COMPLETION");
            dtTranType.Rows.Add("COMPONENT ISSUE");
            dtTranType.Rows.Add("COMPONENT RETURN");
         //   dtTranType.Rows.Add("DECOM2");
         //   dtTranType.Rows.Add("DECOM2 ISSUE");
         //   dtTranType.Rows.Add("DECOM2 RETURN");
         //   dtTranType.Rows.Add("MATERIAL RECEIPT");
        //    dtTranType.Rows.Add("MATERIAL TRANSFER");

            this.ddTranType.DataSource = dtTranType;
            ddTranType.DisplayMember = "Tran Type";
            ddTranType.ValueMember = "Tran Type";


            dtApp = new DataTable();
            dtApp.Columns.Add("APP");
            dtApp.Rows.Add("EPI");
            dtApp.Rows.Add("WB");

            ddApp.DataSource = dtApp;
            ddApp.DisplayMember = "APP";
            ddApp.ValueMember = "APP";

            dtApp.Rows.Add("ALL");

        //    this.comboApp.DataSource = dtApp;
         //   comboApp.DisplayMember = "APP";
        //    comboApp.ValueMember = "APP";

            dtRegion = new DataTable();
            dtRegion.Columns.Add("Region");
            dtRegion.Rows.Add("112");
            dtRegion.Rows.Add("113");

            this.ddRegion.DataSource = dtRegion;


            dtTranSource = new DataTable();
            dtTranSource.Columns.Add("TranSrc");
        //    dtTranSource.Rows.Add("MT");
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
   

        }

        public void getSerialNum(string Ident)
        {
            
            //WebDataGridSerial.DataSource = null;
            string sql = " SELECT [auto_id],status,[serial_num],[parent_serial_num] ,[asset_tag_num],reference,ponum,poline  FROM [MIMDIST].[dbo].[BTS_GoogleIntInvTranSerial]  where uniqueIdent='" + Ident + "'";
            SqlConnection cn = new SqlConnection(constr);
            SqlCommand cmd = new SqlCommand(sql, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet dsSerial = new DataSet();
            da.Fill(dsSerial);
          //  ugBTSSerial.DataSource = dsSerial.Tables[0];                     

        }

        private void ugBTS_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {

          //  e.Layout.Bands[0].Columns["Status"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
            e.Layout.Bands[0].Columns["Status"].ValueList = this.ddStatus;
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

        
        }
        string sql;
        private void loadgrid()
        {
            Cursor.Current = Cursors.WaitCursor;
          //  sql = "select  [auto_id],[status],Replace(replace(Replace(Replace([filename],'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMUS\\outbound\\TEMP\\',''),'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMUS\\outbound\\INVTRAN\\',''),'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMIRE\\outbound\\TEMP\\',''),'\\\\mimsftp\\E$\\sftp\\msgGoogle\\PLATFORM\\PROD\\MIMIRE\\outbound\\INVTRAN\\','') as filename,[fileDate],[ProcessDate],[dekit_summary_id] ,[tran_id],[uniqueIdent] ,[from_loc] ,[to_loc] ,[from_proj],[to_proj],[tran_date],[part_no],[tran_type],[tran_source],[tran_source_num] ,[tran_source_line],[serial_num] ,[qty] ,[reference]  ,[location]  from BTS_GoogleIntInvTran where  uniqueIdent is not null  ";  //,replace(filename,'\\\\mimsftp\\e$\\sftp\\msgGoogle\\PLATFORM\\PROD\\','') as file1

            sql = "SELECT  [MTID]      ,[ProcessDate]      ,[tranxDate]      ,[trxType]      ,[fromLoc]      ,[toLoc]      ,[fromProj]      ,[toProj]      ,[partNo]      ,[qty]      ,[BBNUM]      ,[MTNO]      ,[shipMethod]      ,[comments]      ,[filename]      ,[Status]      ,[int_dep]      ,[MTLineNo]      ,[requestor]      ,[shipTerms]      ,[requestDate]      ,[KanBan_PO_recv]      ,[app]      ,[KBBin]      ,[resent]     ,[EDITranDate]      ,[EDITime]      ,[EDITimeCode]      ,[EDIStatus]      ,[EDICreateDate]      ,[EDIDocument]      ,[EDIICN#]      ,[currencyCode]      ,[Plant]	    FROM [PRODDIST].[dbo].[IntegrationMTRequest]  where MTID is not null";
            if (cmbStatus.Text != "ALL")            {
                sql = sql + " and   status='" + cmbStatus.Value + "'";
            }

            if (txtMTID.Text.Trim() != "")
                sql = sql + " and   [MTID] like '" + txtMTID.Text.Trim() + "%'";

          

            if (this.txtPartNum.Text.Trim() != "")
                sql = sql + " and   [PartNo] like '" + txtPartNum.Text.Trim() + "%'";

            if (this.txtMTNum.Text.Trim() != "")
                sql = sql + " and   [MTNO] like '" + txtMTNum.Text.Trim() + "%'";



            if (this.txtStartdate.Value != "" && this.txtEndDate.Value != "")
            {
                DateTime st_date = Convert.ToDateTime(Convert.ToDateTime(this.txtStartdate.Value).ToShortDateString());
                DateTime end_date = Convert.ToDateTime(Convert.ToDateTime(this.txtEndDate.Value).AddDays(1).ToShortDateString());
                sql = sql + " and   [TranxDate]>='" + st_date + "'  and [TranxDate]<='" + end_date + "'  ";
            }
            else
            {
                MessageBox.Show("Please select start tran date and end tran date");
            }

            sql = sql + "  order by [MTID] desc";
            //   SqlDataSource1.SelectCommand = sql;

            SqlConnection cn = new SqlConnection(constr);
            SqlCommand cmd = new SqlCommand(sql, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            this.grdMTRequest.DataSource = ds.Tables[0];
            this.grdMTRequest.DisplayLayout.Override.CellClickAction = CellClickAction.EditAndSelectText;

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
            DataTable dt = this.grdMTRequest.DataSource as DataTable;
            
            DataRow dr = dt.NewRow();
         //   for (int i = 1; i < ugBTS.Selected.Rows[0].Cells.Count; i++)
           //     dr[i] = ugBTS.Selected.Rows[0].Cells[i].Value;
            dr["mtid"] = 0;
            dt.Rows.Add(dr);
            this.grdMTRequest.DataSource = dt;
            this.grdMTRequest.Rows[this.grdMTRequest.Rows.Count - 1].Activate();

            for (int colnum = 1; colnum < dt.Columns.Count; colnum++)
            {
                this.grdMTRequest.Rows[this.grdMTRequest.Rows.Count - 1].Cells[colnum].Activation = Activation.AllowEdit;
                this.grdMTRequest.Rows[this.grdMTRequest.Rows.Count - 1].Cells[colnum].IgnoreRowColActivation = true;
            }


        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DataTable dt = this.grdMTRequest.DataSource as DataTable;
            DataRow dr = dt.NewRow();
            if (this.grdMTRequest.Selected.Rows.Count > 0)
            {
                dr["mtid"] = 0;
                dr["status"] = "N";
                for (int i = 1; i < this.grdMTRequest.Selected.Rows[0].Cells.Count - 4; i++)

                    if (this.grdMTRequest.Selected.Rows[0].Cells[i].Column.Key.ToUpper() == "STATUS")
                        dr[i] = "N";
                    else if (this.grdMTRequest.Selected.Rows[0].Cells[i].Column.Key.ToUpper() == "UNIQUEIDENT")
                        dr[i] = "";
                    else
                        dr[i] = this.grdMTRequest.Selected.Rows[0].Cells[i].Value;                   
                 
                dt.Rows.Add(dr);
                this.grdMTRequest.DataSource = dt;
                this.grdMTRequest.Rows[this.grdMTRequest.Rows.Count - 1].Activate();
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
            if (this.grdMTRequest.Rows.Count > 0)
            {
                // this.ultraGridExcelExporter1.Export(this.ugBTS, path + "\\GridData.xlsx");
                DataTable dtGrid = grdMTRequest.DataSource as DataTable;
              
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


            int var_MTid=0;
            DateTime var_tranxDate;
            string  var_trxType="";
            string  var_fromLoc="";
            string  var_toLoc="";
            string  var_fromProj="";
            string  var_toProj="";
            string  var_partNo="";
            int     var_qty=0;
            string  var_BBNUM="";
            string  var_MTNO="";
            string  var_shipMethod="";
            string  var_comments="";
            string  var_filename="";
            string  var_Status="";
            string  var_int_dep="";
             int    var_MTLineNo=0;
            string  var_requestor="";
            string  var_shipTerms="";
            DateTime var_requestDate = new DateTime(1900, 1, 1);
            string  var_KanBan_PO_recv="";
            string  var_app="";
            string  var_KBBin="";
            string  var_resent="";
            DateTime var_EDITranDate = new DateTime(1900, 1, 1);
            DateTime var_EDITime = new DateTime(1900, 1, 1);
            string var_EDITimeCode="";
            string var_EDIStatus="";
            DateTime var_EDICreateDate = new DateTime(1900, 1, 1);
            string var_EDIDocument="";
            string var_EDIICN="";
            string var_currencyCode="";
            string var_Plant="";


               
              
             string sqlstr="";
            if (e.Cell.Column.Key == "Save")
            {
                if (e.Cell.Row.Cells["mtid"].Value != "" && e.Cell.Row.Cells["mtid"].Value != null)
                    var_MTid = Convert.ToInt32(e.Cell.Row.Cells["mtid"].Value);
                var_Status = e.Cell.Row.Cells["status"].Value.ToString();
                var_fromLoc = e.Cell.Row.Cells["fromLoc"].Value.ToString();
                var_toLoc = e.Cell.Row.Cells["toLoc"].Value.ToString();
                var_fromProj = e.Cell.Row.Cells["fromProj"].Value.ToString();
                var_toProj = e.Cell.Row.Cells["toProj"].Value.ToString();
                var_tranxDate = Convert.ToDateTime(e.Cell.Row.Cells["tranxDate"].Value.ToString());
                var_partNo = e.Cell.Row.Cells["partNo"].Value.ToString();
                var_trxType = e.Cell.Row.Cells["trxType"].Value.ToString();
                var_BBNUM = e.Cell.Row.Cells["BBNUM"].Value.ToString();
                var_comments = e.Cell.Row.Cells["comments"].Value.ToString();
                var_int_dep = e.Cell.Row.Cells["int_dep"].Value.ToString();
                var_requestor = e.Cell.Row.Cells["requestor"].Value.ToString();

                 var_shipTerms = e.Cell.Row.Cells["shipTerms"].Value.ToString();
                 try
                {
                var_requestDate= Convert.ToDateTime(e.Cell.Row.Cells["requestDate"].Value.ToString());
                }
                 catch { }
                var_KanBan_PO_recv = e.Cell.Row.Cells["KanBan_PO_recv"].Value.ToString();
                var_app = e.Cell.Row.Cells["app"].Value.ToString();
                var_KBBin = e.Cell.Row.Cells["KBBin"].Value.ToString();
                var_resent = e.Cell.Row.Cells["resent"].Value.ToString();
                try
                {
                    var_EDITranDate = Convert.ToDateTime(e.Cell.Row.Cells["EDITranDate"].Value.ToString());
                }
                catch { }
                  try
                {
                var_EDITime = Convert.ToDateTime(e.Cell.Row.Cells["EDITime"].Value.ToString());
                }
                  catch { }
                var_EDITimeCode = e.Cell.Row.Cells["EDITimeCode"].Value.ToString();
                var_EDIStatus = e.Cell.Row.Cells["EDIStatus"].Value.ToString();
                try
                {
                var_EDICreateDate =Convert.ToDateTime( e.Cell.Row.Cells["EDICreateDate"].Value.ToString());
                 }
                catch { }
                var_EDIDocument = e.Cell.Row.Cells["EDIDocument"].Value.ToString();
                var_EDITimeCode = e.Cell.Row.Cells["EDITimeCode"].Value.ToString();
                var_EDIICN = e.Cell.Row.Cells["EDIICN#"].Value.ToString();
                var_currencyCode= e.Cell.Row.Cells["currencyCode"].Value.ToString();
                var_Plant= e.Cell.Row.Cells["Plant"].Value.ToString();
             
                 if (var_MTid == 0)
                    sqlstr = "insert into [IntegrationMTRequest] ([tranxDate]      ,[trxType]      ,[fromLoc]      ,[toLoc]      ,[fromProj]      ,[toProj]      ,[partNo]      ,[qty]      ,[BBNUM]      ,[MTNO]      ,[shipMethod]      ,[comments]            ,[Status]      ,[int_dep]      ,[MTLineNo]      ,[requestor]      ,[shipTerms]      ,[requestDate]      ,[KanBan_PO_recv]      ,[app]      ,[KBBin]      ,[resent]      ,[EDITranDate]      ,[EDITime]      ,[EDITimeCode]      ,[EDIStatus]      ,[EDICreateDate]      ,[EDIDocument]      ,[EDIICN#]      ,[currencyCode]      ,[Plant]) values (  '"+ var_tranxDate + "','" + var_trxType+ "','" + var_fromLoc+ "','" + var_toLoc+ "','" + var_fromProj+ "','" + var_toProj+ "','" + var_partNo+ "','" + var_qty+ "','" + var_BBNUM+ "','" + var_MTNO+ "','" + var_shipMethod+ "','" + var_comments+ "','" + var_Status+ "','" + var_int_dep+ "','" + var_MTLineNo+ "','" + var_requestor+ "','" + var_shipTerms+ "','" + var_requestDate+ "','" + var_KanBan_PO_recv+ "','" + var_app+ "','" + var_KBBin+ "','" + var_resent+ "','" + var_EDITranDate+ "','" + var_EDITime+ "','" + var_EDITimeCode+ "','" + var_EDIStatus+ "','" + var_EDICreateDate+ "','" + var_EDIDocument+ "','" + var_EDIICN+ "','" + var_currencyCode+ "','" + var_Plant+"')";
                 else
                    sqlstr = "update [IntegrationMTRequest] set status='" + var_Status + "' where [MTID]=" + var_MTid;

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
    }
}
