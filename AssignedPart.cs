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
using System.Collections;
using System.Security.Principal;
using Microsoft.Office.Interop.Excel;


namespace Version3
{

    public partial class AssignedPart : Form
    {
        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        public AssignedPart()
        {
            InitializeComponent();
            loadUnassignedGrid();
        }

        public void loadUnassignedGrid()
        {
            string sqlStr = "select  dbo.m_unassigned_parts_A.*  FROM dbo.m_unassigned_parts_A left outer join [dbo].[m_unassigned_part] on m_unassigned_parts_A.part_num=[dbo].[m_unassigned_part].part_num where [dbo].[m_unassigned_part].part_num is null  ";// order by dbo.m_unassigned_parts_A.part_num  ";

            grdUnassigned.Columns.Clear();
            if (txtPartNum.Text.Trim() != "")
                sqlStr = sqlStr + " and dbo.m_unassigned_parts_A.part_num like '" + txtPartNum.Text.Trim() + "%'";
            /// grdUnassigned.DataSource = getDataSet(sqlStr).Tables[0];
            grdUnassigned.AutoGenerateColumns = false;

            grdUnassigned.DataSource = getDataSet(sqlStr).Tables[0];
            if (getDataSet(sqlStr).Tables[0].Rows.Count > 0)
            {
                // grdUnassigned.Width = 
                DataGridViewTextBoxColumn dgvPartNum = new DataGridViewTextBoxColumn();
                dgvPartNum.DataPropertyName = "part_num";
                dgvPartNum.HeaderText = "PART NUM";
                dgvPartNum.Width = 150;

                DataGridViewTextBoxColumn dgvPARTDESC = new DataGridViewTextBoxColumn();
                dgvPARTDESC.DataPropertyName = "description";
                dgvPARTDESC.HeaderText = "PART DESC";
                dgvPARTDESC.Width = 500;

                DataGridViewTextBoxColumn dgvVENDOR = new DataGridViewTextBoxColumn();
                dgvVENDOR.DataPropertyName = "VENDOR";
                dgvVENDOR.HeaderText = "VENDOR";
                dgvVENDOR.Width = 250;

                DataGridViewTextBoxColumn dgvPlatform = new DataGridViewTextBoxColumn();
                dgvPlatform.DataPropertyName = "PLATFORM";
                dgvPlatform.HeaderText = "PLATFORM";
                dgvPlatform.Width = 200;

                DataGridViewTextBoxColumn dgvDUEQTY = new DataGridViewTextBoxColumn();
                dgvDUEQTY.DataPropertyName = "mrp_due_qty";
                dgvDUEQTY.HeaderText = "DUE QTY";
                dgvDUEQTY.Width = 250;
                dgvDUEQTY.Name = "due_qty";

                //DataGridViewTextBoxColumn dgvSupply = new DataGridViewTextBoxColumn();
                //dgvSupply.DataPropertyName = "supply_type";
                //dgvSupply.HeaderText = "SUPPLY TYPE";
                //dgvSupply.Width = 250;


                //    DataGridViewComboBoxColumn dgvSupply = new DataGridViewComboBoxColumn();
                /*  dgvSupply.Items.AddRange(new object[] {
                  "CONSIGNED",
                  "SO-REPORT",
                  "NPI CONSIGNED",
                  "SO NO-REPORT"});
               */
                /*       System.Data.DataTable dtSupplyType = new System.Data.DataTable();
                         DataColumn dcSupplyType1 = new DataColumn();
                         dcSupplyType1.ColumnName = "code";
                         DataColumn dcSupplyType2 = new DataColumn();
                         dcSupplyType2.ColumnName = "desc";
                         dtSupplyType.Columns.Add(dcSupplyType1);
                         dtSupplyType.Columns.Add(dcSupplyType2);
                         DataRow drSupplyType = dtSupplyType.NewRow();
                         drSupplyType[0] = "UNASSIGNED";
                         drSupplyType[1] = "UNASSIGNED";
                         dtSupplyType.Rows.Add(drSupplyType);
                          drSupplyType = dtSupplyType.NewRow();
                         drSupplyType[0] = "CONSIGNED";
                         drSupplyType[1] = "CONSIGNED";
                         dtSupplyType.Rows.Add(drSupplyType);
                         drSupplyType = dtSupplyType.NewRow();
                         drSupplyType[0] = "SO-REPORT";
                         drSupplyType[1] = "SO-REPORT";
                         dtSupplyType.Rows.Add(drSupplyType);
                         drSupplyType = dtSupplyType.NewRow();
                         drSupplyType[0] = "NPI CONSIGNED";
                         drSupplyType[1] = "NPI CONSIGNED";
                         dtSupplyType.Rows.Add(drSupplyType);
                         drSupplyType = dtSupplyType.NewRow();
                         drSupplyType[0] = "SO NO-REPORT";
                         drSupplyType[1] = "SO NO-REPORT";
                         dtSupplyType.Rows.Add(drSupplyType);
             */
                //  dgvSupply.DataSource = dtSupplyType;
                //     dgvSupply.DataPropertyName = "supply_type";
                // dgvSupply.ValueMember = "code";
                //    dgvSupply.DisplayMember = "desc";
                //       dgvSupply.Width = 250;
                //       dgvSupply.HeaderText = "SUPPLY TYPE";

                DataGridViewCheckBoxColumn chkCol = new DataGridViewCheckBoxColumn();
                chkCol.HeaderText = "Select";
                chkCol.Width = 80;
                chkCol.TrueValue = true;


                // chkCol.Selected = true;

                grdUnassigned.Columns.Add(chkCol);
                grdUnassigned.Columns.Add(dgvPartNum);
                grdUnassigned.Columns.Add(dgvPlatform);
                grdUnassigned.Columns.Add(dgvPARTDESC);
                grdUnassigned.Columns.Add(dgvDUEQTY);
                grdUnassigned.Columns.Add(dgvVENDOR);
                //   grdUnassigned.Columns.Add(dgvSupply);

                //  grdUnassigned.Width = this.Width;

                //grdUnassigned.Columns.Add(dgvTLAPartNum);
                //grdUnassigned.Columns.Add(dgvTLAPartDesc);


                // grdUnassigned.Columns[1].DataPropertyName = "PART NUM";
                // grdUnassigned.Columns[2].DataPropertyName = "VENDOR";
                // grdUnassigned.Columns[3].DataPropertyName = "COMMODITY";
                // grdUnassigned.Columns[4].DataPropertyName = "PART DESC";
                // grdUnassigned.Columns[5].DataPropertyName = "SUPPLY TYPE";
                //// grdUnassigned.Columns[5].
                // grdUnassigned.Columns[6].DataPropertyName = "DUE QTY";
                // grdUnassigned.Columns[5].DataPropertyName = "TLA PART NUM";
                // grdUnassigned.Columns[6].DataPropertyName = "TLA DESC";

                //grdUnassigned.AutoGenerateColumns=false

            }
        }

        public DataSet getDataSet(string sqlStr)
        {
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlStr, cn);
            cmd.CommandTimeout = 0;
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;
        }

        //private void btnUpdate_Click(object sender, EventArgs e)
        //{
        //    System.Data.DataTable dtSelected = getallselectPart();
        //    string sql = "";
        //    for (int i = 0; i < dtSelected.Rows.Count; i++)
        //    {
        //        string sqlinsert = " insert into [dbo].[m_unassigned_part] (part_num,platform,create_date,created_by,cur_windchill_type) ";//
        //        string username = WindowsIdentity.GetCurrent().Name.ToString().Replace("MATMOTIONSJ\\", "");
        //        sql = sql + sqlinsert + " values('" + dtSelected.Rows[i]["part_num"].ToString() + "','" + dtSelected.Rows[i]["platform"].ToString() + "','"/* + dtSelected.Rows[i]["VEDNOR"].ToString() + "','" + dtSelected.Rows[i]["COMMODITY"].ToString() + "','" + dtSelected.Rows[i]["TLA PART NUM"].ToString() + "','" + dtSelected.Rows[i]["TLA DESC"].ToString() + "','"*/ + System.DateTime.Now.Date + "','" + username + "','" + dtSelected.Rows[i]["supply_type"] + "')";
        //        //  + dtSelected.Rows[i]["PART NUM"].ToString() + "','"
        //    }
        //    sql = sql + "select @@error";

        //    DataSet ds = getDataSet(sql);
        //    if (ds != null)
        //        if (ds.Tables.Count > 0)
        //            if (ds.Tables[0].Rows[0][0].ToString() == "0")
        //                MessageBox.Show("saved Successfully");


        //}
        private void grdUnassigned_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (grdUnassigned.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == DBNull.Value)
            {
                e.Cancel = true;
            }
        }

        private void dgvAssigned_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (dgvAssignedConfirm.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == DBNull.Value)
            {
                e.Cancel = true;
            }
        }

        //public void getStatus()
        //{
        //    System.Collections.l
        //    ListItemCollection listBoxData = new ListItemCollection();
        //    // Add items to the collection.
        //    listBoxData.Add(new ListItem("apples"));
        //    listBoxData.Add(new ListItem("bananas"));
        //    listBoxData.Add(new ListItem("cherries"));
        //    listBoxData.Add("grapes");
        //    listBoxData.Add("mangos");
        //    listBoxData.Add("oranges");
        //}
        //public void setSupplytype()
        //{

        //    string sql = "select distinct supply_type from nlk_inv_master ";
        //    DataSet dsSupply = getDataSet(sql);
        //    this.cmbSupplyType.DataSource = dsSupply.Tables[0];
        //    cmbSupplyType.DisplayMember="Supply_type";
        //    cmbSupplyType.ValueMember = "Supply_type";
        //}

        public System.Data.DataTable getUnAssignedSelectPart(DataGridView dgv)
        {
            System.Data.DataTable dt = dgv.DataSource as System.Data.DataTable;
            System.Data.DataTable dtSelected = dt.Clone();
            //dt.Select
          //  dtSelected.Columns.Remove("vend_pn");
            dtSelected.Columns.Add("Confirmed Vendor");
            dtSelected.Columns.Add("Lead Time");
            dtSelected.Columns.Add("Price");
            dtSelected.Columns.Add("Confirmed Supply Type");
            dtSelected.Columns.Add("Confirmed By");
            dtSelected.Columns.Add("Comments");
           

            foreach (DataGridViewRow row in dgv.Rows)
            {
                //Get the appropriate cell using index, name or whatever and cast to DataGridViewCheckBoxCell
                DataGridViewCheckBoxCell cell = row.Cells[0] as DataGridViewCheckBoxCell;

                //Compare to the true value because Value isn't boolean
                if (cell.Value == cell.TrueValue)
                {
                    DataRow dr = dtSelected.NewRow();
                    dr["part_num"] = row.Cells[1].Value;                  
                    dr["platform"] = row.Cells[2].Value;
                              
                    dr["vendor"] = row.Cells[5].Value;
                //    dr["supply_type"] = row.Cells[6].Value;
                    dr["description"] = row.Cells[3].Value;
            
                    dr["mrp_due_qty"] = row.Cells[4].Value;
                    dr["Confirmed Vendor"] = "";

                    dr["Lead Time"] = "";
                    dr["Price"] = "";

                    dr["Confirmed Supply Type"] ="";
                    dr["Confirmed By"] = "";

                    dr["Comments"] ="";
         
                   
                   
                    dtSelected.Rows.Add(dr);
                }
                //The value is true
            }           
            return dtSelected;
        }



        public System.Data.DataTable getallselectPart(DataGridView dgv)
        {
            System.Data.DataTable dt = dgv.DataSource as System.Data.DataTable;
            System.Data.DataTable dtSelected = dt.Clone();
            //dt.Select
            //  dtSelected.Columns.Remove("vend_pn");

            foreach (DataGridViewRow row in dgv.Rows)
            {
                //Get the appropriate cell using index, name or whatever and cast to DataGridViewCheckBoxCell
                DataGridViewCheckBoxCell cell = row.Cells[0] as DataGridViewCheckBoxCell;

                //Compare to the true value because Value isn't boolean
                if (cell.Value == cell.TrueValue)
                {
                    DataRow dr = dtSelected.NewRow();
                    dr["part_num"] = row.Cells[1].Value;                   
                 //   dr["cur_windchill_type"] = row.Cells[2].Value;
                 //   if (dgv.Columns.Count > 5)
                 ///   {
                        dr["confirm_by"] = row.Cells[5].Value;
                        dr["Confirm_supply_Type"] = row.Cells[4].Value;
               //     }           
                 
                    dtSelected.Rows.Add(dr);
                }
                //The value is true
            }
            return dtSelected;
        }


        private void button3_Click(object sender, EventArgs e)
        {
            loadUnassignedGrid();
        }

        private void btnLoadGrid_Click(object sender, EventArgs e)
        {

          
             string path = "HKLM\\Software\\Microsoft\\Office\\12.0\\Common\\InstallRoot" + "\\Excel.exe";
             MessageBox.Show(path);  
            
            
       //     dgvAssignedConfirm.Columns.Clear();
        //    string sql = "  Select part_num,cur_windchill_type,case when sent_for_review is null then 'New' when sent_for_review is not null and confirm_date is null then 'Sent For Review'  when sent_for_review is not null and confirm_date is not null and updated_in_mim is null then 'Confirmed Supply Type'  else 'Updated In MIM' end as status,convert(varchar,case when sent_for_review is null then '' when sent_for_review is not null and confirm_supply_type is null then Sent_For_Review  when sent_for_review is not null and confirm_supply_type is not null and  updated_in_mim is null then Confirm_date  else updated_in_mim end,101) as Date,isnull(Confirm_supply_type,'') as Confirm_supply_type,isnull(Confirm_by,'') as Confirm_by,description    from [dbo].[m_unassigned_part]  inner join inv_master on [dbo].[m_unassigned_part].part_num=inv_master.part_no  where isnull(match_windchill,'N')<>'Y'";
        
       //     System.Data.DataTable dtNew= getDataSet(sql).Tables[0].Select("status='New'").CopyToDataTable();   
           

        }

        private void btnUpdateStatus_Click(object sender, EventArgs e)
        {

            string part_num_string = "";
            foreach (DataGridViewRow row in dgvAssignedConfirm.Rows)
            {
                //Get the appropriate cell using index, name or whatever and cast to DataGridViewCheckBoxCell
                DataGridViewCheckBoxCell cell = row.Cells[0] as DataGridViewCheckBoxCell;

                //Compare to the true value because Value isn't boolean
                if (cell.Value == cell.TrueValue)
                {
                    part_num_string = part_num_string + "'" + row.Cells[1].Value + "',";
                }

            }

            string sqlUpdate = " Update [m_unassigned_part] Set ";
            //if (cmbUpdateStatus.SelectedItem.ToString() == "Sent For Review")
            //{
            //    sqlUpdate = sqlUpdate + " sent_for_review='" + System.DateTime.Now.Date + "'";
            //}
            //else if (cmbUpdateStatus.SelectedItem.ToString() == "Confirm For Supply Type")
            //{
            //    sqlUpdate = sqlUpdate + " confirm_date='" + System.DateTime.Now.Date + "' , Confirm_by='" + WindowsIdentity.GetCurrent().Name.ToString().Replace("MATMOTIONSJ\\", "") + "',confirm_supply_type='" + cmbSupplyType.SelectedItem + "'";
            //}
            //else if (cmbUpdateStatus.SelectedItem.ToString() == "Update IN MIM")
            //{
            //    sqlUpdate = sqlUpdate + " update_in_mim='" + System.DateTime.Now.Date + "'";
            //}
            //else if (cmbUpdateStatus.SelectedItem.ToString() == "Match Windchill")
            //{
            //    if (chkMatchWindChill.Checked)
            //        sqlUpdate = sqlUpdate + " match_windchill='Y'";
            //    else
            //        sqlUpdate = sqlUpdate + " match_windchill='N'";
            //}

            sqlUpdate = sqlUpdate + " where part_num in (" + part_num_string.Substring(0, part_num_string.Length - 1) + ")   select @@error";
            try
            {
                SqlConnection cn = new SqlConnection(constr);
                SqlCommand cmd = new SqlCommand(sqlUpdate, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                if (ds != null)
                    if (ds.Tables.Count > 0)
                        if (ds.Tables[0].Rows[0][0].ToString() == "0")
                            MessageBox.Show("Status has been updated successfully");

            }
            catch (Exception exp)
            {
                MessageBox.Show("Error to update");
            }


        }

        //private void cmbUpdateStatus_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (cmbUpdateStatus.SelectedItem.ToString() == "Confirm For Supply Type")
        //    {
        //        lblConfirmBy.Visible = true;
        //        cmbConfirmBy.Visible = true;
        //        lblSupplyType.Visible = true;
        //        cmbSupplyType.Visible = true;
        //        chkMatchWindChill.Visible = false;
        //    }
        //    else if (cmbUpdateStatus.SelectedItem.ToString() == "Match Windchill")
        //    {
        //        lblConfirmBy.Visible = false;
        //        cmbConfirmBy.Visible = false;
        //        lblSupplyType.Visible = false;
        //        chkMatchWindChill.Visible = true;
        //        cmbSupplyType.Visible = false;

        //    }
        //    else
        //    {
        //        lblConfirmBy.Visible = false;
        //        cmbConfirmBy.Visible = false;
        //        lblSupplyType.Visible = false;
        //        cmbSupplyType.Visible = false;
        //        chkMatchWindChill.Visible = false;
        //    }
        //}

    

        string filename = "";
    

    //    private void button1_Click(object sender, EventArgs e)
    //    {

    //        System.Data.DataTable dt = this.dgvAssignedReview.DataSource as System.Data.DataTable;
    //        System.Data.DataTable dtSelected = dt.Clone();
    //        //dt.Select
          

    //        foreach (DataGridViewRow row in dgvAssignedReview.Rows)
    //        {
    //            //Get the appropriate cell using index, name or whatever and cast to DataGridViewCheckBoxCell
    //            DataGridViewCheckBoxCell cell = row.Cells[0] as DataGridViewCheckBoxCell;

    //            //Compare to the true value because Value isn't boolean
    //            if (cell.Value == cell.TrueValue)
    //            {
    //                DataRow dr = dtSelected.NewRow();
    //                //for (int i = 0; i < dtSelected.Columns.Count - 1; i++)
    //              //  {
    //                    dr["part_num"] = row.Cells[1].Value;
    //                    dr["cur_windchill_type"] = row.Cells[2].Value;
    //                    dr["date"] = row.Cells[4].Value;
    //                    dr["status"] = row.Cells[3].Value;
    //                    dr["confirm_supply_type"] = row.Cells[5].Value;
    //                    dr["confirm_by"] = row.Cells[6].Value;
    //                    dr["description"] = row.Cells[7].Value;
    //              //  }
                                         
                   
    //                dtSelected.Rows.Add(dr);
    //            }
    //       }

    //        if (dtSelected.Rows.Count > 0)
    //        {
    //            filename = "AssignedPartFile" + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Year.ToString() + '_' + System.DateTime.Now.Hour + "_" + System.DateTime.Now.Minute + "_" + System.DateTime.Now.Second + ".xml";
    //            filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + filename;
    //            dtSelected.WriteXml(filename);
    //            String EXL = "C:\\Program Files\\Microsoft Office\\Office14\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
    //            System.Diagnostics.Process proc = new System.Diagnostics.Process();
    //            proc.StartInfo.FileName = EXL;
    //            proc.StartInfo.Arguments = filename;
    //            proc.StartInfo.UseShellExecute = false;
    //            proc.StartInfo.RedirectStandardOutput = true;
    //            proc.Start();
    //        }
    //        else
    //            MessageBox.Show("Before export, you need to check the parts you want to export");
            
           
    //}

    



        public void saveData()
        {
            System.Data.DataTable dtSelected = this.getUnAssignedSelectPart(grdUnassigned);
            string sql = "";
            for (int i = 0; i < dtSelected.Rows.Count; i++)
            {
                string sqlinsert = " insert into [dbo].[m_unassigned_part] (part_num,platform,create_date,sent_for_review,created_by,cur_windchill_type) ";
                string username = WindowsIdentity.GetCurrent().Name.ToString().Replace("MATMOTIONSJ\\", "");
                sql = sql + sqlinsert + " values('" + dtSelected.Rows[i]["part_num"].ToString() + "','" + dtSelected.Rows[i]["platform"].ToString() + "','"/* + dtSelected.Rows[i]["VEDNOR"].ToString() + "','" + dtSelected.Rows[i]["COMMODITY"].ToString() + "','" + dtSelected.Rows[i]["TLA PART NUM"].ToString() + "','" + dtSelected.Rows[i]["TLA DESC"].ToString() + "','"*/ + System.DateTime.Now.Date + "','" + System.DateTime.Now.Date + "','" + username + "','SO-REPORT')";// +dtSelected.Rows[i]["supply_type"] + "')";
                //  + dtSelected.Rows[i]["PART NUM"].ToString() + "','"
            }
            sql = sql + "select @@error";

            DataSet ds = getDataSet(sql);
            if (ds != null)
                if (ds.Tables.Count > 0)
                    if (ds.Tables[0].Rows[0][0].ToString() == "0")
                        MessageBox.Show("Saved Successfully");
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

            saveData();
            filename = "AssignedPartFile" + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Year.ToString() + '_' + System.DateTime.Now.Hour + "_" + System.DateTime.Now.Minute + "_" + System.DateTime.Now.Second + ".xlsx";
            filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + filename;
            if (System.IO.File.Exists(filename))
            {
                System.IO.File.Delete(filename);
            }
            if (!System.IO.File.Exists(filename))
            {
                System.IO.FileStream fs = System.IO.File.Create(filename);
            }



            System.Data.DataTable dtSelectedPart = getUnAssignedSelectPart(grdUnassigned);

          


            if (dtSelectedPart.Rows.Count > 0)
            {
                filename = "AssignedPartFile" + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Year.ToString() + '_' + System.DateTime.Now.Hour + "_" + System.DateTime.Now.Minute + "_" + System.DateTime.Now.Second + ".xml";
                filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + filename;
                dtSelectedPart.WriteXml(filename);
                String EXL ="HKLM\\Software\\Microsoft\\Office\\12.0\\Common\\InstallRoot"+"\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = EXL;
                proc.StartInfo.Arguments = filename;
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.Start();
            }
            else
                MessageBox.Show("Before export, you need to check the parts you want to export");


         // GenerateExcelFile(dtSelectedPart,  filename );
        }

        private void grdUnassigned_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgvAssigned_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (dgvAssignedConfirm.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewLinkCell)
                {
                    if (dgvAssignedConfirm.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "Update")
                    {
                        string part_num = dgvAssignedConfirm.Rows[e.RowIndex].Cells[1].Value.ToString();
                        string confirmed_Supply_type = dgvAssignedConfirm.Rows[e.RowIndex].Cells[3].Value.ToString();
                        string Confirmed_by = dgvAssignedConfirm.Rows[e.RowIndex].Cells[4].Value.ToString();
                        string status= dgvAssignedConfirm.Rows[e.RowIndex].Cells[6].Value.ToString();
                        //string part_num = dgvAssigned.Rows[e.RowIndex].Cells[1].Value.ToString();
                        UpdatePart(part_num, confirmed_Supply_type, status, Confirmed_by);
                    }
                }
            }
        }
        private void UpdatePart(string part_num, string confirmed_Supply_type, string status,string Confirmed_by)
       {
              string sqlUpdate = " Update [m_unassigned_part] Set ";
              if (status == "Sent For Review")
            {
                sqlUpdate = sqlUpdate + " sent_for_review='" + System.DateTime.Now.Date + "'";
            }
              else if (status == "Confirmed Supply Type")
            {
                sqlUpdate = sqlUpdate + " confirm_date='" + System.DateTime.Now.Date + "' , Confirm_by='" + Confirmed_by + "',confirm_supply_type='" + confirmed_Supply_type + "'";
            }
              else if (status == "Updated in MIM")
            {
                sqlUpdate = sqlUpdate + " updated_in_mim='" + System.DateTime.Now.Date + "'";
            }
           // else if (cmbUpdateStatus.SelectedItem.ToString() == "Match Windchill")
         //   {
              //  if (chkMatchWindChill.Checked)
               //     sqlUpdate = sqlUpdate + " match_windchill='Y'";
              //  else
               //     sqlUpdate = sqlUpdate + " match_windchill='N'";
         //   }

            sqlUpdate = sqlUpdate + " where part_num ='" + part_num + "'   select @@error";
            try
            {
                SqlConnection cn = new SqlConnection(constr);
                SqlCommand cmd = new SqlCommand(sqlUpdate, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                if (ds != null)
                    if (ds.Tables.Count > 0)
                        if (ds.Tables[0].Rows[0][0].ToString() == "0")
                            MessageBox.Show("Status has been updated successfully");

            }
            catch (Exception exp)
            {
                MessageBox.Show("Error to update");
            }

       }

        private void btnSentForReview_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtSelected = getallselectPart(dgvAssignedConfirm);
            string part_num_str="";

            for (int i = 0; i < dtSelected.Rows.Count; i++)
                part_num_str = part_num_str +"'"+ dtSelected.Rows[i]["part_num"].ToString() +"', ";


            string sqlUpdate = " Update [m_unassigned_part] Set  sent_for_review='" + System.DateTime.Now.Date + "'";


            sqlUpdate = sqlUpdate + " where part_num  in (" + part_num_str.Substring(0,part_num_str.Length-2) + ")    select @@error";
            try
            {
                SqlConnection cn = new SqlConnection(constr);
                SqlCommand cmd = new SqlCommand(sqlUpdate, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                if (ds != null)
                    if (ds.Tables.Count > 0)
                        if (ds.Tables[0].Rows[0][0].ToString() == "0")
                            MessageBox.Show("Status has been updated successfully");

            }
            catch (Exception exp)
            {
                MessageBox.Show("Error to update");
            }

        }

        private void btnSaveConfType_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dtSelected = dgvAssignedConfirm.DataSource as System.Data.DataTable;
            string sqlUpdate = "";
            bool AllRight = true;

            for (int i = 0; i < dtSelected.Rows.Count; i++)
            {
                if ( dtSelected.Rows[i]["confirm_Supply_type"].ToString() !="")
                {
                    if (dtSelected.Rows[i]["confirm_by"].ToString() != "")
                    {
                        sqlUpdate = sqlUpdate + "  Update [m_unassigned_part] Set confirm_date='" + System.DateTime.Now.Date + "' , Confirm_by='" + dtSelected.Rows[i]["confirm_by"].ToString() + "',confirm_supply_type='" + dtSelected.Rows[i]["confirm_Supply_type"].ToString() + "'";
                        sqlUpdate = sqlUpdate + "  where part_num = '" + dtSelected.Rows[i]["part_num"].ToString() + "' ";
                        DataGridViewCell cell = dgvAssignedConfirm.Rows[i].Cells[4] as DataGridViewCell;
                        cell.Style.BackColor = System.Drawing.Color.White;
                    }
                    else
                    {
                        DataGridViewCell cell = dgvAssignedConfirm.Rows[i].Cells[4] as DataGridViewCell;
                        cell.Style.BackColor = System.Drawing.Color.Red;
                        AllRight = false;
                    }
                }
            }
            if (AllRight == false)
            {
                MessageBox.Show("Please select the Confirm by");

            }
            else
            {
                try
                {
                    sqlUpdate = sqlUpdate + "  select @@error ";
                    SqlConnection cn = new SqlConnection(constr);
                    SqlCommand cmd = new SqlCommand(sqlUpdate, cn);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    if (ds != null)
                        if (ds.Tables.Count > 0)
                            if (ds.Tables[0].Rows[0][0].ToString() == "0")
                                MessageBox.Show("Status has been updated successfully");

                }
                catch (Exception exp)
                {
                    MessageBox.Show("Error to update");
                }
            }
            refreshGrid();
        }

        private void btnSaveINMIM_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtSelected = getallselectPart(this.dgvAssignedUpdMIM);
            string part_num_str = "";

            for (int i = 0; i < dtSelected.Rows.Count; i++)
                part_num_str = part_num_str + "'" + dtSelected.Rows[i]["part_num"].ToString() + "', ";


            string sqlUpdate = " Update [m_unassigned_part] Set  updated_in_mim='" + System.DateTime.Now.Date + "'";
            sqlUpdate = sqlUpdate + " where part_num in (" + part_num_str.Substring(0,part_num_str.Length-2) + ")   select @@error";

            try
            {
                SqlConnection cn = new SqlConnection(constr);
                SqlCommand cmd = new SqlCommand(sqlUpdate, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                if (ds != null)
                    if (ds.Tables.Count > 0)
                        if (ds.Tables[0].Rows[0][0].ToString() == "0")
                            MessageBox.Show("Status has been updated successfully");

            }
            catch (Exception exp)
            {
                MessageBox.Show("Error to update");
            }



        }

        private void tabControl1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
                loadUnassignedGrid();
            else
            {
                string sql = "  Select part_num,case when sent_for_review is null then 'New' when sent_for_review is not null and confirm_date is null then 'Sent For Review'  when sent_for_review is not null and confirm_date is not null and updated_in_mim is null then 'Confirmed Supply Type'  else 'Updated In MIM' end as status,convert(varchar,case when sent_for_review is null then create_date when sent_for_review is not null and confirm_supply_type is null then Sent_For_Review  when sent_for_review is not null and confirm_supply_type is not null and  updated_in_mim is null then Confirm_date  else updated_in_mim end,101) as Date,isnull(Confirm_supply_type,'') as Confirm_supply_type,isnull(Confirm_by,'') as Confirm_by,description    from [dbo].[m_unassigned_part]  inner join inv_master on [dbo].[m_unassigned_part].part_num=inv_master.part_no  where isnull(match_windchill,'N')<>'Y'";
                string status = "";
                System.Data.DataSet dsAssignedPart = getDataSet(sql);
                //if (tabControl1.SelectedIndex == 1)
                //{
                //    status = "New";  ,cur_windchill_type
                //    System.Data.DataTable dtNew = dsAssignedPart.Tables[0].Select("status='"+status+"'").CopyToDataTable();
                //    loadAssignedPart(status, dtNew, this.dgvAssignedReview);

                //}
                if (tabControl1.SelectedIndex == 1)
                {
                    status = "Sent For Review";
                    if (dsAssignedPart.Tables[0].Select("status='" + status + "'").Length > 0)
                    {
                        System.Data.DataTable dtNew = dsAssignedPart.Tables[0].Select("status='" + status + "'").CopyToDataTable();
                        loadAssignedPart(status, dtNew, this.dgvAssignedConfirm);
                    }
                    
                }
                else if (tabControl1.SelectedIndex == 2)
                {
                    status = "Confirmed Supply Type";
                    if (dsAssignedPart.Tables[0].Select("status='" + status + "'").Length > 0)
                    {
                        System.Data.DataTable dtNew = dsAssignedPart.Tables[0].Select("status='" + status + "'").CopyToDataTable();
                        loadAssignedPart(status, dtNew, this.dgvAssignedUpdMIM);
                    }
                }
                else if (tabControl1.SelectedIndex == 3)
                {
                    status = "Updated In MIM";
                  
                    if (dsAssignedPart.Tables[0].Select("status='" + status + "'").Length > 0)
                    {
                        System.Data.DataTable dtNew = dsAssignedPart.Tables[0].Select("status='" + status + "'").CopyToDataTable();
                        loadAssignedPart(status, dtNew, this.dgvAssignedMatchWind);
                    }
                }
              
              
            }

        }


        public void refreshGrid()
        {

            string sql = "  Select part_num,case when sent_for_review is null then 'New' when sent_for_review is not null and confirm_date is null then 'Sent For Review'  when sent_for_review is not null and confirm_date is not null and updated_in_mim is null then 'Confirmed Supply Type'  else 'Updated In MIM' end as status,convert(varchar,case when sent_for_review is null then create_date when sent_for_review is not null and confirm_supply_type is null then Sent_For_Review  when sent_for_review is not null and confirm_supply_type is not null and  updated_in_mim is null then Confirm_date  else updated_in_mim end,101) as Date,isnull(Confirm_supply_type,'') as Confirm_supply_type,isnull(Confirm_by,'') as Confirm_by,description    from [dbo].[m_unassigned_part]  inner join inv_master on [dbo].[m_unassigned_part].part_num=inv_master.part_no  where isnull(match_windchill,'N')<>'Y'";
            string status = "";
            System.Data.DataSet dsAssignedPart = getDataSet(sql);
            //if (tabControl1.SelectedIndex == 1)
            //{
            //    status = "New";  ,cur_windchill_type
            //    System.Data.DataTable dtNew = dsAssignedPart.Tables[0].Select("status='"+status+"'").CopyToDataTable();
            //    loadAssignedPart(status, dtNew, this.dgvAssignedReview);

            //}
            if (tabControl1.SelectedIndex == 1)
            {
                status = "Sent For Review";
                if (dsAssignedPart.Tables[0].Select("status='" + status + "'").Length > 0)
                {
                    System.Data.DataTable dtNew = dsAssignedPart.Tables[0].Select("status='" + status + "'").CopyToDataTable();
                    loadAssignedPart(status, dtNew, this.dgvAssignedConfirm);
                }

            }
            else if (tabControl1.SelectedIndex == 2)
            {
                status = "Confirmed Supply Type";
                if (dsAssignedPart.Tables[0].Select("status='" + status + "'").Length > 0)
                {
                    System.Data.DataTable dtNew = dsAssignedPart.Tables[0].Select("status='" + status + "'").CopyToDataTable();
                    loadAssignedPart(status, dtNew, this.dgvAssignedUpdMIM);
                }
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                status = "Updated In MIM";

                if (dsAssignedPart.Tables[0].Select("status='" + status + "'").Length > 0)
                {
                    System.Data.DataTable dtNew = dsAssignedPart.Tables[0].Select("status='" + status + "'").CopyToDataTable();
                    loadAssignedPart(status, dtNew, this.dgvAssignedMatchWind);
                }
            }
        }

        public void loadAssignedPart(string status,System.Data.DataTable dt,DataGridView dgv)
        {
            dgv.Columns.Clear();
            //if (status=="New")
            //{
            //    dgv.Columns.Clear();
            //}
            //else if (status=="Sent For Review")
            //{
            //    this.dgv.Columns.Clear();
            //}
            //else if (status=="Confirmed Supply Type")
            //{
            //    this.dgv.Columns.Clear();
            //}
            dgv.AutoGenerateColumns = false;
            DataGridViewTextBoxColumn dgvPartNum = new DataGridViewTextBoxColumn();
            dgvPartNum.Name = "part_num";
            dgvPartNum.DataPropertyName = "part_num";
            dgvPartNum.HeaderText = "Part Num";
            dgvPartNum.Width = 100;

         /* DataGridViewTextBoxColumn dgvWindchill = new DataGridViewTextBoxColumn();
            dgvWindchill.Name = "cur_windchill_type";
            dgvWindchill.DataPropertyName = "cur_windchill_type";
            dgvWindchill.HeaderText = "Curr Windchill Type";
            dgvWindchill.Width = 200;
*/

            DataGridViewTextBoxColumn dgvdesc = new DataGridViewTextBoxColumn();
            dgvdesc.Name = "desc";
            dgvdesc.DataPropertyName = "description";
            dgvdesc.HeaderText = "Part Description";
            dgvdesc.Width = 400;


            DataGridViewTextBoxColumn dgvStatus = new DataGridViewTextBoxColumn();
            dgvStatus.Name = "Status";
            dgvStatus.DataPropertyName = "status";
            dgvStatus.HeaderText = "Status";
            dgvStatus.Width = 200;                          
           
            DataGridViewComboBoxColumn dgvSupply = new DataGridViewComboBoxColumn();
            dgvSupply.Name = "Supply_type";
            System.Data.DataTable dtSupplyType = new System.Data.DataTable();

            DataColumn dcSupplyType1 = new DataColumn();
            dcSupplyType1.ColumnName = "code";

            DataColumn dcSupplyType2 = new DataColumn();
            dcSupplyType2.ColumnName = "desc";

            dtSupplyType.Columns.Add(dcSupplyType1);
            dtSupplyType.Columns.Add(dcSupplyType2);
          
            
            DataRow drSupplyType = dtSupplyType.NewRow();


            drSupplyType = dtSupplyType.NewRow();
            drSupplyType[0] = "";
            drSupplyType[1] = "";
            dtSupplyType.Rows.Add(drSupplyType);

            drSupplyType = dtSupplyType.NewRow();
            drSupplyType[0] = "CONSIGNED";
            drSupplyType[1] = "CONSIGNED";
            dtSupplyType.Rows.Add(drSupplyType);

            drSupplyType = dtSupplyType.NewRow();
            drSupplyType[0] = "SO-REPORT";
            drSupplyType[1] = "SO-REPORT";
            dtSupplyType.Rows.Add(drSupplyType);

            drSupplyType = dtSupplyType.NewRow();
            drSupplyType[0] = "NPI CONSIGNED";
            drSupplyType[1] = "NPI CONSIGNED";
            dtSupplyType.Rows.Add(drSupplyType);

            drSupplyType = dtSupplyType.NewRow();
            drSupplyType[0] = "SO NO-REPORT";
            drSupplyType[1] = "SO NO-REPORT";
            dtSupplyType.Rows.Add(drSupplyType);

            dgvSupply.DataSource = dtSupplyType;
            dgvSupply.DataPropertyName = "Confirm_supply_type";
            dgvSupply.ValueMember = "code";
            dgvSupply.DisplayMember = "desc";
            dgvSupply.Width = 350;
            dgvSupply.HeaderText = "Confirmed Supply Type";


          

/*
            DataGridViewComboBoxColumn dgvStatus = new DataGridViewComboBoxColumn();
            dgvStatus.Name = "Status";
            System.Data.DataTable dtStatus = new System.Data.DataTable();
            DataColumn dc1 = new DataColumn();
            dc1.ColumnName = "code";
            DataColumn dc2 = new DataColumn();
            dc2.ColumnName = "desc";
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);
            DataRow dr = dt.NewRow();
            dr[0] = "Sent For Review";
            dr[1] = "Sent For Review";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr[0] = "Confirmed Supply Type";
            dr[1] = "Confirmed Supply Type";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr[0] = "Updated in MIM";
            dr[1] = "Updated in MIM";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr[0] = "Match Windchill";
            dr[1] = "Match Windchill";
            dt.Rows.Add(dr);


            // dgvStatus.Items.AddRange(new object[] {
            //  "Sent For Review","Confirmed For Supply Type","Updated in MIM","Match Windchill"});

            dgvStatus.DataSource = dt;
            dgvStatus.DisplayMember = "desc";
            dgvStatus.ValueMember = "code";

            dgvStatus.HeaderText = "Status";
            dgvStatus.Width = 200;
            dgvStatus.DataPropertyName = "Status";
*/
            DataGridViewTextBoxColumn dgvDate = new DataGridViewTextBoxColumn();
            // dgvStatus.Items.AddRange(new object[] {
            //     "Confirm For Supply Type","Pending for Update in MIM","Pending for Matching","Pending For Review"}

            //);
            dgvDate.DataPropertyName = "Date";
            dgvDate.HeaderText = "Date";
            dgvDate.Width = 100;

          

            DataGridViewComboBoxColumn dgvConfirmBy = new DataGridViewComboBoxColumn();
            System.Data.DataTable dtConfirmBy = new System.Data.DataTable();
            DataColumn dcConfirmBy1 = new DataColumn();
            dcConfirmBy1.ColumnName = "code";
            DataColumn dcConfirmBy2 = new DataColumn();
            dcConfirmBy2.ColumnName = "desc";
            dtConfirmBy.Columns.Add(dcConfirmBy1);
            dtConfirmBy.Columns.Add(dcConfirmBy2);
            DataRow drConfirmBy = dtConfirmBy.NewRow();


            drConfirmBy = dtConfirmBy.NewRow();
            drConfirmBy[0] = "";
            drConfirmBy[1] = "";
            dtConfirmBy.Rows.Add(drConfirmBy);
          //  drConfirmBy = dtConfirmBy.NewRow();

            drConfirmBy = dtConfirmBy.NewRow();
            drConfirmBy[0] = "Zach";
            drConfirmBy[1] = "Zach";
            dtConfirmBy.Rows.Add(drConfirmBy);
            drConfirmBy = dtConfirmBy.NewRow();

            drConfirmBy[0] = "Venessa";
            drConfirmBy[1] = "Venessa";
            dtConfirmBy.Rows.Add(drConfirmBy);

            drConfirmBy = dtConfirmBy.NewRow();
            drConfirmBy[0] = "Jensen";
            drConfirmBy[1] = "Jensen";
            dtConfirmBy.Rows.Add(drConfirmBy);

            drConfirmBy = dtConfirmBy.NewRow();
            drConfirmBy[0] = "Raghav";
            drConfirmBy[1] = "Raghav";
            dtConfirmBy.Rows.Add(drConfirmBy);

            drConfirmBy = dtConfirmBy.NewRow();
            drConfirmBy[0] = "NPI";
            drConfirmBy[1] = "NPI";
            dtConfirmBy.Rows.Add(drConfirmBy);

            dgvConfirmBy.DataSource = dtConfirmBy;
            dgvConfirmBy.DataPropertyName = "Confirm_by";
            dgvConfirmBy.ValueMember = "code";
            dgvConfirmBy.DisplayMember = "desc";
            dgvConfirmBy.Width = 200;
            dgvConfirmBy.HeaderText = "Confirm By";

            DataGridViewCheckBoxColumn chkAssignCol = new DataGridViewCheckBoxColumn();
            chkAssignCol.HeaderText = "Select";
            chkAssignCol.Width = 80;
            chkAssignCol.TrueValue = true;


            dgvPartNum.ReadOnly = true;
            dgvDate.ReadOnly = true;
            dgvSupply.ReadOnly = true;
            dgvConfirmBy.ReadOnly = true;
            dgvStatus.ReadOnly = true;
            dgvdesc.ReadOnly = true;


            dgv.Columns.Add(chkAssignCol);
            dgv.Columns.Add(dgvPartNum);
           // dgv.Columns.Add(dgvWindchill);
            dgv.Columns.Add(dgvDate);
           
             dgv.Columns.Add(dgvSupply);//dgvSupply
             dgv.Columns.Add(dgvConfirmBy);
           

            dgv.Columns.Add(dgvStatus);
            dgv.Columns.Add(dgvdesc);
            if (status == "Sent For Review")
            {
                chkAssignCol.Visible = false;
                dgvSupply.ReadOnly = false;
                dgvConfirmBy.ReadOnly = false;
            }
            dgv.DataSource = dt;



        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
        //private void button5_Click(object sender, EventArgs e)
        //{
        

        //}

    
        private void btnExportUpdMIM_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = this.dgvAssignedUpdMIM.DataSource as System.Data.DataTable;
            System.Data.DataTable dtSelected = dt.Clone();
            //dt.Select

            dtSelected.TableName = "exportTable";
            foreach (DataGridViewRow row in dgvAssignedUpdMIM.Rows)
            {
                //Get the appropriate cell using index, name or whatever and cast to DataGridViewCheckBoxCell
                DataGridViewCheckBoxCell cell = row.Cells[0] as DataGridViewCheckBoxCell;

                //Compare to the true value because Value isn't boolean
                if (cell.Value == cell.TrueValue)
                {
                    DataRow dr = dtSelected.NewRow();
                    //for (int i = 0; i < dtSelected.Columns.Count - 1; i++)
                    //  {
                    dr["part_num"] = row.Cells[1].Value;
                   // dr["cur_windchill_type"] = row.Cells[2].Value;
                    dr["date"] = row.Cells[2].Value;
                    dr["status"] = row.Cells[5].Value;
                    dr["confirm_supply_type"] = row.Cells[3].Value;
                    dr["confirm_by"] = row.Cells[4].Value;
                    dr["description"] = row.Cells[6].Value;
                    //  }


                    dtSelected.Rows.Add(dr);
                }
            }

            if (dtSelected.Rows.Count > 0)
            {
                filename = "AssignedPartFile" + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Year.ToString() + '_' + System.DateTime.Now.Hour + "_" + System.DateTime.Now.Minute + "_" + System.DateTime.Now.Second + ".xml";
                filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + filename;
                dtSelected.WriteXml(filename);
                Microsoft.Office.Interop.Word.Application appVersion = new Microsoft.Office.Interop.Word.Application();
              //  MessageBox.Show(appVersion.Path);
                String EXL = appVersion.Path + "\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = EXL;
                proc.StartInfo.Arguments = filename;
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.Start();
            }
            else
                MessageBox.Show("Before export, you need to check the parts you want to export");
        }

        private void btnExportConfirm_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtSelected = this.dgvAssignedConfirm.DataSource as System.Data.DataTable;
           
         
            //foreach (DataGridViewRow row in dgvAssignedUpdMIM.Rows)
            //{
            //    //Get the appropriate cell using index, name or whatever and cast to DataGridViewCheckBoxCell
            //    DataGridViewCheckBoxCell cell = row.Cells[0] as DataGridViewCheckBoxCell;

            //    //Compare to the true value because Value isn't boolean
            //    if (cell.Value == cell.TrueValue)
            //    {
            //        DataRow dr = dtSelected.NewRow();
            //        //for (int i = 0; i < dtSelected.Columns.Count - 1; i++)
            //        //  {
            //        dr["part_num"] = row.Cells[1].Value;
            //        dr["cur_windchill_type"] = row.Cells[2].Value;
            //        dr["date"] = row.Cells[3].Value;
            //        dr["status"] = row.Cells[6].Value;
            //        dr["confirm_supply_type"] = row.Cells[4].Value;
            //        dr["confirm_by"] = row.Cells[5].Value;
            //        dr["description"] = row.Cells[7].Value;
            //        //  }


            //        dtSelected.Rows.Add(dr);
            //    }
            //}
            dtSelected.TableName = "exportTable";
            if (dtSelected.Rows.Count > 0)
            {
                filename = "AssignedPartFile" + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Year.ToString() + '_' + System.DateTime.Now.Hour + "_" + System.DateTime.Now.Minute + "_" + System.DateTime.Now.Second + ".xml";
                filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + filename;
                dtSelected.WriteXml(filename);
                try
                {
                    String EXL = "C:\\Program Files\\Microsoft Office\\Office14\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.StartInfo.FileName = EXL;
                    proc.StartInfo.Arguments = filename;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.Start();
                }
                catch
                {
                    String EXL = "C:\\Program Files (x86)\\Microsoft Office\\Office14\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.StartInfo.FileName = EXL;
                    proc.StartInfo.Arguments = filename;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.Start();
                }

            }
            else
                MessageBox.Show("Before export, you need to check the parts you want to export");
        }

        private void btnExportReview_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = this.dgvAssignedConfirm.DataSource as System.Data.DataTable;
            System.Data.DataTable dtSelected = dt.Clone();
            //dt.Select


            foreach (DataGridViewRow row in dgvAssignedConfirm.Rows)
            {
                //Get the appropriate cell using index, name or whatever and cast to DataGridViewCheckBoxCell
                DataGridViewCheckBoxCell cell = row.Cells[0] as DataGridViewCheckBoxCell;

                //Compare to the true value because Value isn't boolean
                if (cell.Value == cell.TrueValue)
                {
                    DataRow dr = dtSelected.NewRow();
                    //for (int i = 0; i < dtSelected.Columns.Count - 1; i++)
                    //  {
                    dr["part_num"] = row.Cells[1].Value;
                  //  dr["cur_windchill_type"] = row.Cells[2].Value;
                    dr["date"] = row.Cells[3].Value;
                    dr["status"] = row.Cells[4].Value;
                    dr["confirm_supply_type"] = row.Cells[5].Value;
                    dr["confirm_by"] = row.Cells[6].Value;
                   dr["description"] = row.Cells[5].Value;
                    //  }
                    dtSelected.Rows.Add(dr);
                }
            }
            dtSelected.TableName = "exportTable";
            if (dtSelected.Rows.Count > 0)
            {
                filename = "AssignedPartFile" + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Year.ToString() + '_' + System.DateTime.Now.Hour + "_" + System.DateTime.Now.Minute + "_" + System.DateTime.Now.Second + ".xml";
                filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + filename;
                dtSelected.WriteXml(filename);
                try
                {
                    String EXL = "C:\\Program Files\\Microsoft Office\\Office14\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.StartInfo.FileName = EXL;
                    proc.StartInfo.Arguments = filename;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.Start();
                }
                catch
                {
                    String EXL = "C:\\Program Files (x86)\\Microsoft Office\\Office14\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.StartInfo.FileName = EXL;
                    proc.StartInfo.Arguments = filename;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.Start();
                }
            }
            else
                MessageBox.Show("Before export, you need to check the parts you want to export");

        }

        private void btnMatchWindchill_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dtSelected = getallselectPart(dgvAssignedMatchWind);
            string part_num_str = "";

            for (int i = 0; i < dtSelected.Rows.Count; i++)
                part_num_str = part_num_str + "'" + dtSelected.Rows[i]["part_num"].ToString() + "', ";


            string sqlUpdate = " Update [m_unassigned_part] Set ";

              if (chkMatchWindChill.Checked)
                    sqlUpdate = sqlUpdate + " match_windchill='Y'";
                else
                    sqlUpdate = sqlUpdate + " match_windchill='N'";


              sqlUpdate = sqlUpdate + " where part_num in (" + part_num_str.Substring(0, part_num_str.Length - 2) + ")   select @@error";
            try
            {
                SqlConnection cn = new SqlConnection(constr);
                SqlCommand cmd = new SqlCommand(sqlUpdate, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                if (ds != null)
                    if (ds.Tables.Count > 0)
                        if (ds.Tables[0].Rows[0][0].ToString() == "0")
                            MessageBox.Show("Status has been updated successfully");

            }
            catch (Exception exp)
            {
                MessageBox.Show("Error to update");
            }
        }

        private void btnExportWindchill_Click(object sender, EventArgs e)
        {


            System.Data.DataTable dt = this.dgvAssignedMatchWind.DataSource as System.Data.DataTable;
            System.Data.DataTable dtSelected = dt.Clone();
            //dt.Select

            dtSelected.TableName = "exportTable";
            foreach (DataGridViewRow row in dgvAssignedMatchWind.Rows)
            {
                //Get the appropriate cell using index, name or whatever and cast to DataGridViewCheckBoxCell
                DataGridViewCheckBoxCell cell = row.Cells[0] as DataGridViewCheckBoxCell;

                //Compare to the true value because Value isn't boolean
                if (cell.Value == cell.TrueValue)
                {
                    DataRow dr = dtSelected.NewRow();
                    //for (int i = 0; i < dtSelected.Columns.Count - 1; i++)
                    //  {
                    dr["part_num"] = row.Cells[1].Value;
                  //  dr["cur_windchill_type"] = row.Cells[2].Value;
                    dr["date"] = row.Cells[2].Value;
                    dr["status"] = row.Cells[5].Value;
                    dr["confirm_supply_type"] = row.Cells[3].Value;
                    dr["confirm_by"] = row.Cells[4].Value;
                    dr["description"] = row.Cells[6].Value;
                    //  }


                    dtSelected.Rows.Add(dr);
                }
            }

            if (dtSelected.Rows.Count > 0)
            {
                filename = "AssignedPartFile" + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Year.ToString() + '_' + System.DateTime.Now.Hour + "_" + System.DateTime.Now.Minute + "_" + System.DateTime.Now.Second + ".xml";
                filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + filename;
                dtSelected.WriteXml(filename);
                try
                {
                    String EXL = "C:\\Program Files\\Microsoft Office\\Office14\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.StartInfo.FileName = EXL;
                    proc.StartInfo.Arguments = filename;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.Start();
                }
                catch
                {
                    String EXL = "C:\\Program Files (x86)\\Microsoft Office\\Office14\\Excel.exe";  /// PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.StartInfo.FileName = EXL;
                    proc.StartInfo.Arguments = filename;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.Start();
                }

            }
            else
                MessageBox.Show("Before export, you need to check the parts you want to export");
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string sql = "  delete from m_unassigned_part where part_num in ( select m_unassigned_part.part_num from m_unassigned_part left outer join m_unassigned_parts_A on m_unassigned_part.part_num=m_unassigned_parts_A.part_num where m_unassigned_parts_A.part_num is null) and sent_for_review is not null and confirm_date is null  select @@error ";
            try
            {
                SqlConnection cn = new SqlConnection(constr);
                SqlCommand cmd = new SqlCommand(sql, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                if (ds != null)
                    if (ds.Tables.Count > 0)
                        if (ds.Tables[0].Rows[0][0].ToString() == "0")
                            MessageBox.Show("Parts have been synced successfully");

            }
            catch (Exception exp)
            {
                MessageBox.Show("Error to update");
            }
            refreshGrid();
        }


        //public void GenerateExcelFile(System.Data.DataTable dt, string paramFileFullPath)
        //{
        //    // Global missing reference for objects we are not defining...
        //    object missing = System.Reflection.Missing.Value;

        //    // Create an Excel object and add workbook...
        //    ApplicationClass excel = new ApplicationClass();
        //    Workbook workbook = excel.Application.Workbooks.Add(true);

        //    //Rename the first/default sheet name
        //    if (dt.TableName.ToString().Trim() != string.Empty)
        //    {
        //        Worksheet ws = (Worksheet)excel.Worksheets.get_Item(1);
        //        ws.Name = dt.TableName.ToString();
        //    }

        //    int iCol = 0;
        //    // Add column headings...
        //    // if (_IsHeaderIncluded == true)
        //    // {
        //    foreach (DataColumn c in dt.Columns)
        //    {
        //        iCol++;
        //        excel.Cells[1, iCol] = c.ColumnName.Replace('_',' ').ToUpper();
        //    }
        //    //  }

        //    // for each row of data...
        //    int iRow = 1;
        //    foreach (DataRow r in dt.Rows)
        //    {
        //        iRow++;
        //        // add each row's cell data...
        //        iCol = 0;
        //        foreach (DataColumn c in dt.Columns)
        //        {
        //            iCol++;
        //            //   if (_IsHeaderIncluded == true)
        //            //   {
        //            excel.Cells[iRow, iCol] = r[c.ColumnName];
        //            //  }
        //            //   else
        //            //  {
        //            //      excel.Cells[iRow, iCol] = r[c.ColumnName];
        //            //  }
        //        }
        //    }

        //    // If wanting to make Excel visible and activate the worksheet...
        //    excel.Visible = true;
        //    Worksheet worksheet = (Worksheet)excel.ActiveSheet;
        //    ((_Worksheet)worksheet).Activate();

        //    // If wanting excel to shutdown...
        //    //((_Application)excel).Quit();
        //}

        ////public void CreateExcel(string filename)
        ////{
        ////    //  excelWorkbook.SaveAs("E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\"+filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
        ////    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        ////    //string workbookPath = "E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\VendorSheet.xls";

        ////        filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + filename;


        ////    if (System.IO.File.Exists(filename))
        ////        System.IO.File.Delete(filename);


        ////    string workbookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\NPI_MRP_TEMPLATE.xlsx";
        ////    Microsoft.Office.Interop.Excel.Workbook excelWorkbook1 = excelApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        ////    try
        ////    {
        ////        excelWorkbook1.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
        ////    }
        ////    catch (Exception e)
        ////    {
        ////        string msg = e.Message.ToString();
        ////    }
        ////    finally
        ////    {
        ////        //Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
        ////        excelWorkbook1.Close(true, false, Type.Missing);
        ////        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook1);
        ////        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        ////    }

        ////}

        //public void export_datagridview_to_excel(DataGridView dgv)
        //{
        //    int cols;

        //    filename = "AssignedPartFile" + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Year.ToString() + '_' + System.DateTime.Now.Hour + "_" + System.DateTime.Now.Minute + "_" + System.DateTime.Now.Second + ".csv";
        //    filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + filename;
        //    if (System.IO.File.Exists(filename))
        //    {
        //        System.IO.File.Delete(filename);
        //    }
        //    if (!System.IO.File.Exists(filename))
        //    {
        //        System.IO.FileStream fs = System.IO.File.Create(filename);
        //        fs.Close();
        //    }

        //    //open file
        //    System.IO.StreamWriter wr = new System.IO.StreamWriter(filename);
        //    //determine the number of columns and write columns to file
        //    cols = dgv.Columns.Count;
        //    for (int i = 0; i < cols; i++)
        //    {
        //        DataGridViewColumn column=dgv.Columns[0];
        //        wr.Write(column.Name.ToString().ToUpper() + "\t");
        //    }
        //    wr.WriteLine();
        //    //write rows to excel file
        //    for (int i = 0; i < (dgv.Rows.Count - 1); i++)
        //    {
        //        for (int j = 0; j < cols; j++)
        //        {


        //            DataGridViewCheckBoxCell cell = dgv.Rows[i].Cells[0] as DataGridViewCheckBoxCell;

        //            //Compare to the true value because Value isn't boolean
        //            if (cell.Value == cell.TrueValue)

        //            {
        //                if (dgv.Rows[i].Cells[j].Value != null)
        //                    wr.Write(dgv.Rows[i].Cells[j].Value + "\t");
        //                else
        //                {
        //                    wr.Write("\t");
        //                }
        //            }

        //        }
        //        wr.WriteLine();
        //}
        //    //close file
        //    wr.Close();
        //}
    }
}
