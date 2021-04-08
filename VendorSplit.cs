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
using Infragistics.Win.UltraWinEditors;
using Infragistics.Win.UltraWinGrid;
using Microsoft.VisualBasic.PowerPacks;


namespace Version3
{
    public partial class VendorSplit : Form
    {
        public VendorSplit()
        {
            InitializeComponent();
            bindVendor();
            bindQrter();
            bindSite();
            bindQrterCopy();
            bindVendorCurSplit();
        }

        private void VendorSplit_Load(object sender, EventArgs e)
        {
         
         
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
            string strqry = " SELECT distinct sku_code as part_no FROM v_vendor_split_parts_list  order by sku_code ";
            DataSet ds = getDataSet(strqry, constr);
        //    cmbPartNum.DataSource = ds.Tables[0];
          //  cmbPartNum.ValueMember = "vendor_code";
        //    cmbPartNum.DisplayMember = "vendor_code";

            this.cmbPartNum.DataSource = ds.Tables[0];
            cmbPartNum.ValueMember = "part_no";
            cmbPartNum.DisplayMember = "part_no";

         //   this.comboVendor.DataSource = ds.Tables[0];
          //  comboVendor.ValueMember = "vendor_code";
           // comboVendor.DisplayMember = "vendor_code";
        }


        public void bindVendorCurSplit()
        {
            string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
            string strqry = " select vendor_code,address_name from dbo.apmaster where status_type = 5 and vend_class_code = 'VENDOR' AND vendor_code NOT LIKE 'DECOM%' AND flag_1099 = 1 ";
            DataSet ds = getDataSet(strqry, constr);
          
            this.cmbVendorCurSplit.DataSource = ds.Tables[0];
            cmbVendorCurSplit.ValueMember = "vendor_code";
            cmbVendorCurSplit.DisplayMember = "vendor_code";

        }

        public void bindQrter()
        {

            DataTable dtQuarter = new DataTable();
            dtQuarter.Columns.Add(new DataColumn("Quarter"));
            DataRow dr = dtQuarter.NewRow();
            dr[0] = "2014Q01";
            dtQuarter.Rows.Add(dr);

            dr = dtQuarter.NewRow();
            dr[0] = "2014Q02";
            dtQuarter.Rows.Add(dr);

            dr = dtQuarter.NewRow();
            dr[0] = "2014Q03";
            dtQuarter.Rows.Add(dr);

            dr = dtQuarter.NewRow();
            dr[0] = "2014Q04";
            dtQuarter.Rows.Add(dr);


            dr = dtQuarter.NewRow();
            dr[0] = "2015Q01";
            dtQuarter.Rows.Add(dr);

            dr = dtQuarter.NewRow();
            dr[0] = "2015Q02";
            dtQuarter.Rows.Add(dr);

            dr = dtQuarter.NewRow();
            dr[0] = "2015Q03";
            dtQuarter.Rows.Add(dr);

            dr = dtQuarter.NewRow();
            dr[0] = "2015Q04";
            dtQuarter.Rows.Add(dr);
           

         
       }

        public void bindSite()
        {

            DataTable dtSite= new DataTable();
            dtSite.Columns.Add(new DataColumn("Site"));
            DataRow dr = dtSite.NewRow();
            dr[0] = "AT";
            dtSite.Rows.Add(dr);
            dr = dtSite.NewRow();
            dr[0] = "NL";
            dtSite.Rows.Add(dr);
         

            cmbSite.DataSource = dtSite;
            cmbSite.ValueMember = "Site";
            cmbSite.DisplayMember = "Site";

            cmbSiteCurSplit.DataSource = dtSite;
            cmbSiteCurSplit.ValueMember = "Site";
            cmbSiteCurSplit.DisplayMember = "Site";


        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            try
            {
                //string vendorVal = (item.Controls["comboVendor"] as ComboBox).Text.ToString();
                //DataRow dr1 = dsGrid.Tables[0].NewRow();
                //dr1["site"] = cmbSite.SelectedText.ToString();
                //dr1["vendor_code"] = vendorVal;
                //dr1["sku_code"] = cmbPartNum.SelectedText;
                //dr1["quarter"] = ComboQtr.SelectedText;
                //dr1["split"] = (item.Controls["txtSplit"] as TextBox).Text;

                //dsGrid.Tables[0].Rows.Add(dr1);
                //dsGrid.AcceptChanges();
                //repeaterVendorSplit.DataSource = dsGrid.Tables[0];
                repeaterVendorSplit.AddNew();
            }
            catch (Exception exp)
            {
            }
         //   m_vendor_split_tableBindingSource.AddNew();
            
            //repeaterVendorSplit.AddNew();
        //  dsGrid.AcceptChanges();
         //   m_vendor_split_tableBindingSource.DataSource = dsGrid.Tables[0];
         //   repeaterVendorSplit.DataSource = m_vendor_split_tableBindingSource;
        /*    DataRepeaterItem item = repeaterVendorSplit.CurrentItem;
            if ((item.Controls["comboVendor"] as ComboBox).Text != "SELECT")// &&  (item.Controls["txtSplit"] as TextBox).Text !="0.00" ))
            {
                string vendorVal = (item.Controls["comboVendor"] as ComboBox).Text.ToString();
                DataRow[] dr = dsGrid.Tables[0].Select("vendor_code='" + vendorVal + "'");
                if (dr.Length == 1)
                {
                    TextBox txt = item.Controls["txtSplit"] as TextBox;
                    if (txt != null)
                    {
                        dr[0]["Split"] = (item.Controls["txtSplit"] as TextBox).Text.ToString();
                        dsGrid.Tables[0].AcceptChanges();
                    }
                }
                else
                {
                    DataRow dr1 = dsGrid.Tables[0].NewRow();
                    dr1["site"] = cmbSite.SelectedText.ToString();
                    dr1["vendor_code"] = vendorVal;
                    dr1["sku_code"] = cmbPartNum.SelectedText;
                    dr1["quarter"] = ComboQtr.SelectedText;
                    dr1["split"] = (item.Controls["txtSplit"] as TextBox).Text;

                    dsGrid.Tables[0].Rows.Add(dr1);
                    dsGrid.AcceptChanges();
                }

            }
            try
            {
               // repeaterVendorSplit.AddNew();
                DataRow dr = dsGrid.Tables[0].NewRow();
                dr["site"] = cmbSite.SelectedText.ToString();
                dr["vendor_code"] = "SELECT";
                dr["sku_code"] = cmbPartNum.SelectedText;
                dr["quarter"] = ComboQtr.SelectedText;
                dr["split"] = "0.00";            

                dsGrid.Tables[0].Rows.Add(dr);
                dsGrid.AcceptChanges();
                repeaterVendorSplit.DataSource = dsGrid.Tables[0];
            }
            catch
            {

            }*/
        }
        DataTable dtRepeater;
        DataSet dsGrid;
        bool setFilterFlag = false;
        private void btnFilter_Click(object sender, EventArgs e)
        {
            setFilterFlag = true;
            bindgrid();           
        }

        private void bindgrid()
        {
            
            try
            {
                if (cmbPartNum.Text == "" || cmbSite.Text == "" || ComboQtr.Text == "")
                {
                    MessageBox.Show("Please select part num, site and quarter");
                }
                else
                {
                    string sqlGrid = "SELECT [site],[sku_code],[vendor_code],[split],[quarter],comp_pn   FROM [dbo].[m_vendor_split_table] where sku_code='" + cmbPartNum.Text + "' and site='" + cmbSite.Text + "' and  quarter='" + ComboQtr.Text + "'     select * from  v_vendor_splt_display   where [PART NUMBER]='" + cmbPartNum.Text + "'   select distinct quarter from  v_vendor_splt_display   where [PART NUMBER]='" + cmbPartNum.Text + "'  and  QUARTER>=(SELECT     cast(YEAR(getdate()) AS varchar) + 'Q' + replicate('0', (2 - len(DATENAME(Quarter, GETDATE())))) + DATENAME(Quarter, GETDATE()) ) ";
                    string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
                    dsGrid = getDataSet(sqlGrid, constr);
                    dtRepeater = getDataSet(sqlGrid, constr).Tables[0];
                  
                

              // dtRepeater= dtRepeater.Select("sku_code='" + cmbPartNum.Text + "' and site='" + cmbSite.Text + "' and  quarter='" + ComboQtr.Text + "'").CopyToDataTable();
                m_vendor_split_tableBindingSource.DataSource = dtRepeater;
               // m_vendor_split_tableBindingSource.Filter = "sku_code='" + cmbPartNum.Text + "' and site='" + cmbSite.Text + "' and  quarter='" + ComboQtr.Text + "'";
               repeaterVendorSplit.DataSource = dtRepeater;
               
                grdVendorSplit.DataSource = dsGrid.Tables[1];

                grdQuarterCombo.DataSource = null;
                DataTable dt = dsGrid.Tables[2];
                DataRow dr = dt.NewRow();
                dr["quarter"] = "ALL";
                dt.Rows.InsertAt(dr,0);

                grdQuarterCombo.DataSource = dt;
                grdQuarterCombo.DisplayMember = "quarter";
                grdQuarterCombo.ValueMember = "quarter";
                grdQuarterCombo.SelectedRow = grdQuarterCombo.Rows[0];
              //  dsGrid.Clear();
                //dsGrid.Tables.Add(dtRepeater);
             //   repeaterVendorSplit.CurrentItemIndex = repeaterVendorSplit.ItemCount - 1;
                }
            }
            catch(Exception e)
            {
                MessageBox.Show("Error to display data");
            }
        }


        private void ultraNumericEditor1_BeforeEnterEditMode(object sender, CancelEventArgs e)
        {

        }

        private void ultraNumericEditor1_ValidationError(object sender, Infragistics.Win.UltraWinEditors.ValidationErrorEventArgs e)
        {
           
        }

        private void repeaterVendorSplit_CurrentItemIndexChanged(object sender, EventArgs e)
        {

            //dsGrid.AcceptChanges();
           // m_vendor_split_tableBindingSource.DataSource = dsGrid.Tables[0];
          //  updateTotal();
        }
        DataTable dtVendor;
        DataSet dsRepeater;
        DataTable dtCompPart;
        private void repeaterVendorSplit_ItemCloned(object sender, Microsoft.VisualBasic.PowerPacks.DataRepeaterItemEventArgs e)
        {
            string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
            try
            {
                UltraCombo CmbVend = e.DataRepeaterItem.Controls["comboVendor"] as UltraCombo;
               string vend = cmbPartNum.SelectedText;
              // dsRepeater = getDataSet("SELECT distinct vendor_code FROM [MIMDIST].[dbo].[apmaster] where status_type=5 and [vend_class_code]='VENDOR'    ", constr);
               dsRepeater = getDataSet("select vendor_code, address_name from dbo.apmaster where status_type = 5 and vend_class_code = 'VENDOR' AND vendor_code NOT LIKE 'DECOM%' AND flag_1099 = 1    ", constr);
               dtVendor = dsRepeater.Tables[0];
                DataRow dr = dtVendor.NewRow();
                dr[0] = "SELECT";
                
                dtVendor.Rows.InsertAt(dr, 0);
                if (CmbVend.DataSource == null)
                {
                    CmbVend.DataSource = dtVendor;
                    CmbVend.DisplayMember = "vendor_code";
                    CmbVend.ValueMember = "vendor_code";
                }
            //   Cmb.SelectedIndexChanged += new EventHandler(comboVendor_SelectedIndexChanged);
            //    Cmb.SelectionChangeCommitted += new EventHandler(comboBox1_SelectionChangeCommitted);
                UltraNumericEditor txtsplit = e.DataRepeaterItem.Controls["txtSplit"] as UltraNumericEditor;
             //   txtsplit.Leave += new EventHandler(txtSplit_Leave);
                txtsplit.AfterExitEditMode += new EventHandler(txtSplit_AfterExitEditMode);
               // txtsplit.ValueChanged += new CancelEventHandler(txtSplit_BeforeEnterEditMode);
            }

            catch (Exception exp1) { MessageBox.Show(exp1.Message.ToString()); }

            try
            {
                UltraCombo Cmb = e.DataRepeaterItem.Controls["cmpCompPart"] as UltraCombo;
               // dsRepeater = getDataSet(" select sku_code as part_no from  v_vendor_split_parts_list  ", constr);
                dsRepeater = getDataSet(" select sku_code as part_no from  v_vendor_split_parts_list  ", constr);
                dtCompPart = dsRepeater.Tables[0];
                DataRow dr = dtCompPart.NewRow();
                dr[0] = "SELECT";

                dtCompPart.Rows.InsertAt(dr, 0);
                if (Cmb.DataSource == null)
                {
                    Cmb.DataSource = dtCompPart;
                    Cmb.DisplayMember = "part_no";
                    Cmb.ValueMember = "part_no";
                }
             //   if (lblPlanType.Text == "OPTION")
               // {
                    Cmb.Visible = true;                  
                    Infragistics.Win.Misc.UltraLabel lbl = e.DataRepeaterItem.Controls["ultraLabel3"] as Infragistics.Win.Misc.UltraLabel;
                    lbl.Visible = true;                    
              //  }


            //    Cmb.SelectedIndexChanged += new EventHandler(combocom_SelectedIndexChanged);
              //  UltraNumericEditor txtsplit = e.DataRepeaterItem.Controls["txtSplit"] as UltraNumericEditor;
                //   txtsplit.Leave += new EventHandler(txtSplit_Leave);
             //   txtsplit.AfterExitEditMode += new EventHandler(txtSplit_AfterExitEditMode);
                // txtsplit.ValueChanged += new CancelEventHandler(txtSplit_BeforeEnterEditMode);
            }

            catch (Exception exp1) { MessageBox.Show(exp1.Message.ToString()); }

        }
        bool selectIndexFlag = false;
        private void comboVendor_SelectedIndexChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    ComboBox Cmb = repeaterVendorSplit.CurrentItem.Controls["comboVendor"] as ComboBox;
            //    for (int i = 0; i < dtVendor.Rows.Count; i++)
            //    {

            //        if (dtVendor.Rows[i][0].ToString() == Cmb.Text.ToString())
            //            Cmb.SelectedIndex = i;

            //    }
            //}
            //catch { }
         
         
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {

            ComboBox senderComboBox = (ComboBox)sender;

            // Change the length of the text box depending on what the user has  
            // selected and committed using the SelectionLength property. 
            //if (senderComboBox.SelectionLength > 0)
            //{

            //    senderComboBox.SelectedValue = senderComboBox.Items[senderComboBox.SelectedIndex].ToString();
            //}
        }

        private void AddUpdateDataTable()
        {
           
        }

        private void repeaterVendorSplit_DrawItem(object sender, Microsoft.VisualBasic.PowerPacks.DataRepeaterItemEventArgs e)
        {

           
                try
                {
                    // UltraNumericEditor txtsplit = e.DataRepeaterItem.Controls["txtSplit"] as UltraNumericEditor;
                    //  txtsplit.ValueChanged += new EventHandler(txtSplit_ValueChanged);
                    // txtsplit.BeforeEnterEditMode += new CancelEventHandler(txtSplit_BeforeEnterEditMode);
                    UltraCombo Cmb = e.DataRepeaterItem.Controls["comboVendor"] as UltraCombo;
                   // if (Cmb.SelectedIndex == 0)
                   // {
                    if (Cmb.SelectedRow == null)
                    {
                        if (dtVendor != null)
                        {
                            for (int i = 0; i < dtVendor.Rows.Count; i++)
                            {
                                if (this.dtRepeater.Rows.Count > e.DataRepeaterItem.ItemIndex)
                                    if (dtVendor.Rows[i][0].ToString() == dtRepeater.Rows[e.DataRepeaterItem.ItemIndex]["vendor_code"].ToString())
                                       
                                Cmb.SelectedRow = Cmb.Rows[i];
                                // Cmb.SelectedIndex = i;
                            }
                        }
                    }
                    else
                    {
                        /* 
                         
                          for (int i = 0; i < dtVendor.Rows.Count; i++)
                           {
                               if (this.dtRepeater.Rows.Count > e.DataRepeaterItem.ItemIndex)
                                   if (dtVendor.Rows[i][0].ToString() == Cmb.Text.ToString())
                                       Cmb.SelectedIndex = i;

                           }
                         
                         */

                    }
                }
                catch (Exception exp)
                {
                    MessageBox.Show(exp.Message.ToString());
                }


                try
                {

                    UltraCombo Cmb = e.DataRepeaterItem.Controls["cmpCompPart"] as UltraCombo;
                    if (Cmb.SelectedRow == null)
                    {
                        if (this.dtCompPart != null)
                        {
                            if (cmbPartNum.SelectedRow != null)
                            {
                                for (int i = 0; i < dtCompPart.Rows.Count; i++)
                                {
                                    if (this.dtRepeater.Rows.Count > e.DataRepeaterItem.ItemIndex)
                                        if (dtCompPart.Rows[i][0].ToString() == dtRepeater.Rows[e.DataRepeaterItem.ItemIndex]["comp_pn"].ToString())
                                            Cmb.SelectedRow = Cmb.Rows[i];
                                    //Cmb.SelectedText = dtRepeater.Rows[e.DataRepeaterItem.ItemIndex]["comp_pn"].ToString();

                                }
                            }
                        }
                    }
                }
                catch (Exception exp)
                {
                    MessageBox.Show(exp.Message.ToString());
                }
              
            
        }




        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            try
            {
                //DeleteLine();
               // repeaterVendorSplit.RemoveAt(repeaterVendorSplit.CurrentItemIndex);
                //   DataTable datatable = repeaterVendorSplit.DataSource as DataTable;dsRepeater.Tables[0];// 
                dsGrid.Tables[0].Rows[repeaterVendorSplit.CurrentItemIndex].Delete();
                dsGrid.AcceptChanges();
                dtRepeater = dsGrid.Tables[0];
                m_vendor_split_tableBindingSource.DataSource = dtRepeater;
                repeaterVendorSplit.DataSource = dtRepeater;
                //repeaterVendorSplit.DataSource = dsGrid.Tables[0];//.Select("[sku_code]='" + cmbPartNum.Text + "' and [quarter]='" + ComboQtr.Text + "' and [site]='" + cmbSite.SelectedText.ToString() + "'").CopyToDataTable();
            }
            catch(Exception exp2)
            { }
        }

        public void DeleteLine()
        {
          //  if (repeaterVendorSplit.CurrentItem.Controls["comb"].Text.Trim() != "")
           // {
                string sql = "delete from [m_vendor_split_table] where sku_code='" + this.cmbPartNum.Text + "' and quarter='" + this.ComboQtr.Text + "' and site='" + cmbSite.Text.ToString() + "' and vendor_code='" + this.repeaterVendorSplit.CurrentItem.Controls["comboVendor"].Text + "'  select @@error";
                string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
                DataSet ds = getDataSet(sql, constr);
                if (ds.Tables[0].Rows[0][0].ToString() == "0")
                {
                    MessageBox.Show("Item Deleted");
                    bindgrid();
                }
                else
                {
                    MessageBox.Show("Error to delete Item");
                }
           // }
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            updateTotal();

            if (txtTotalSplit.Value.ToString() != "0")
            {
                if (Convert.ToDecimal(txtTotalSplit.Value.ToString()) < Convert.ToDecimal(100.5) && Convert.ToDecimal(txtTotalSplit.Value.ToString()) > Convert.ToDecimal(99.5))
                {
                    string part_num = this.cmbPartNum.Text;
                    string site = this.cmbSite.Text;
                    string quarter = this.ComboQtr.Text;
                    string sql = "";
                    string sqlDelete = "delete from [MIMDIST].[dbo].[m_vendor_split_table] where sku_code='" + part_num + "' and site='" + site + "' and quarter='" + quarter + "'";
                    bool duplicateRec = false;
                    bool EmptyVendor = false;
                    string vendor_list = "";
                    string comp_pn = cmbPartNum.Text;
                    bool compSame = false;
                    bool blankComp = false;
                    bool InvalidVendorCode = false;

                    for (int i = 0; i < repeaterVendorSplit.ItemCount; i++)
                    {
                        repeaterVendorSplit.CurrentItemIndex = i;
                        Microsoft.VisualBasic.PowerPacks.DataRepeaterItem item = repeaterVendorSplit.CurrentItem;
                        Decimal split = 0;
                        string vendor = "";

                        if (repeaterVendorSplit.CurrentItem.Controls["txtSplit"].Text.ToString() != "")
                            split = Convert.ToDecimal(repeaterVendorSplit.CurrentItem.Controls["txtSplit"].Text.ToString());

                        if (repeaterVendorSplit.CurrentItem.Controls["comboVendor"].Text.ToString() != "SELECT")
                            vendor = repeaterVendorSplit.CurrentItem.Controls["comboVendor"].Text.ToString();
                        if (dtVendor.Select("vendor_code='" + vendor + "'").Length == 0)
                        {
                            InvalidVendorCode = true;
                           
                            repeaterVendorSplit.CurrentItem.Controls["comboVendor"].Focus();
                            sql = "";
                            break;
                            
                        }

                     //   if (this.lblPlanType.Text == "OPTION")
                     //   {
                            comp_pn = "";
                            comp_pn = repeaterVendorSplit.CurrentItem.Controls["cmpCompPart"].Text.ToString();
                            if (comp_pn == part_num)
                            {
                                compSame = true;
                            }
                            else if (comp_pn == "" || comp_pn == "SELECT")
                            {
                                blankComp = true;

                            }
                     //   }

                        if (vendor_list.Contains(vendor +"-" + comp_pn))
                        {
                            duplicateRec = true;
                        }
                        if (vendor == "")
                        {
                            EmptyVendor = true;
                        }

                        vendor_list = vendor_list + "," + vendor +"-" + comp_pn;
                        if (part_num != "" && site != "" && quarter != "")// && split != 0.00)
                        {
                            sql = sql + " insert into [MIMDIST].[dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],comp_pn,[split],[quarter]) values ('" + site + "','" + part_num + "','" + vendor + "','" + comp_pn + "','" + split + "','" + quarter + "')";
                        }

                    }
                    if (EmptyVendor)
                    {
                        MessageBox.Show("Select the vendor");
                    }
                    else if (InvalidVendorCode)
                         MessageBox.Show("Vendor Code is not valid.Please check");
                  //  else if (compSame)
                  //      MessageBox.Show("For OPTION Part, component part can not be same as part number. Please select different Component part number","Error Message",MessageBoxButtons.OK);
                    else if (blankComp)
                        MessageBox.Show("For OPTION Part, There must be component part number", "Error Message", MessageBoxButtons.OK);

                    else
                    {
                        if (!duplicateRec)
                        {
                            if (lblErrorMessage.Text.Trim() == "")
                            {
                                sql = sqlDelete + sql;
                                string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;

                                SqlConnection cn = new SqlConnection(constr);
                                cn.Open();
                                SqlTransaction tran = cn.BeginTransaction("Transaction");
                                SqlCommand cmd = new SqlCommand(sql, cn);
                                cmd.Transaction = tran;

                                try
                                {
                                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                                    DataSet ds = new DataSet();
                                    da.Fill(ds);
                                    tran.Commit();
                                    MessageBox.Show("Updated successfully");
                                    bindgrid();

                                }
                                catch (System.Exception exp)
                                {
                                    string error = exp.Message.ToString();
                                    tran.Rollback();
                                    MessageBox.Show("Failed to Update:" + error);
                                }
                                finally
                                {
                                    cn.Close();
                                }

                            }
                            else
                            {
                                MessageBox.Show(lblErrorMessage.Text, "Error Message", MessageBoxButtons.OK);
                                bindgrid();
                            }
                        }
                        else
                        {
                            MessageBox.Show(" Duplicate Vendor");
                        }

                    }
                }
                else
                {
                    MessageBox.Show("Split should be 100%.Please check the splits");
                }

            }
            else
            {

                if (repeaterVendorSplit.ItemCount == 0)
                {
                    string part_num = this.cmbPartNum.Text;
                    string site = this.cmbSite.Text;
                    string quarter = this.ComboQtr.Text;
                    string sqlDelete = "delete from [MIMDIST].[dbo].[m_vendor_split_table] where sku_code='" + part_num + "' and site='" + site + "' and quarter='" + quarter + "'";
                    string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;

                    SqlConnection cn = new SqlConnection(constr);
                    cn.Open();

                    SqlCommand cmd = new SqlCommand(sqlDelete, cn);


                    try
                    {
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataSet ds = new DataSet();
                        da.Fill(ds);

                        MessageBox.Show("Updated successfully");
                        bindgrid();

                    }
                    catch (System.Exception exp)
                    {
                        string error = exp.Message.ToString();

                        MessageBox.Show("Failed to Update:" + error);
                    }
                    finally
                    {
                        cn.Close();
                    }


                }
            }

         
            }


        public void clearALL()
        {
            this.cmbPartNum.Text = "";
            this.cmbSite.Text = "";
            this.ComboQtr.Text="";
            this.repeaterVendorSplit.DataSource = null;
        }
        private void txtSplit_Leave(object sender, EventArgs e)
        {
           // updateTotal();
          //  getTotal();

           
        }

        public void getTotal()
        {
            Decimal total = 0;
            Decimal currSplit = 0;
            DataTable dt = dsGrid.Tables[0].Copy();
              dt=  repeaterVendorSplit.DataSource as DataTable;

              foreach (DataRow dr in dt.Rows)
            {
                currSplit = Convert.ToDecimal(dr["split"].ToString());
                total = total + currSplit;

            }
            txtTotalSplit.Value = total;

        }
        private void txtSplit_textChanged(object sender, EventArgs e)
        {

           try
            {
                UltraNumericEditor txt = sender as UltraNumericEditor;
               DataRepeaterItem item = (txt.Parent) as DataRepeaterItem;
               item.Controls["ultraLabel3"].Text = txt.Value.ToString();

            }
            catch (Exception e1)
            {
            }

        }
        private void updateTotal()
        {
            decimal split=0;
            DataSet dsRepeaterCopy = dsRepeater.Clone();
            DataTable dtRepeater1 = dsRepeaterCopy.Tables[0];
           dtRepeater1= repeaterVendorSplit.DataSource as DataTable;
          /* foreach (DataRepeaterItem item in repeaterVendorSplit.ItemTemplate)
           {
               split = split + Convert.ToDecimal(item.Controls["txtSplit"].Text.ToString());
           }*/
           for (int i = 0; i < repeaterVendorSplit.ItemCount; i++)
            {
                repeaterVendorSplit.CurrentItemIndex=i;
                DataRepeaterItem item = repeaterVendorSplit.CurrentItem;
                split = split + Convert.ToDecimal(item.Controls["txtSplit"].Text.ToString());
            }
            txtTotalSplit.Value=split;            
        }

        private void ultraLabel3_Click(object sender, EventArgs e)
        {

        }

        private void ultraLabel3_TextChanged(object sender, EventArgs e)
        {
          //  updateTotal();
        }

        private void txtSplit_BeforeEnterEditMode(object sender, CancelEventArgs e)
        {
           // UltraNumericEditor txt = sender as UltraNumericEditor;
           // txt.SelectAll();
        }

        private void txtSplit_ValidationError(object sender, ValidationErrorEventArgs e)
        {
           // UltraNumericEditor txt = sender as UltraNumericEditor;
            //e.RetainFocus = false;
           // txt.Value = 0;
        }

        private void m_vendor_split_tableBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void txtSplit_AfterExitEditMode(object sender, EventArgs e)
        {
         //   repeaterVendorSplit.DataSource = dsGrid.Tables[0];
        }
        DataTable dtCopy;
        private void btnCopy_Click(object sender, EventArgs e)
        {
            dtCopy = repeaterVendorSplit.DataSource as DataTable;

            string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
            string qrySave = "";
            foreach (DataRow dr in dtCopy.Rows)
            {

             
             //   qrySave = qrySave + " Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('" + dr["SITE"] + "','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + ComboQtrCopy.SelectedText.ToString() + "','" + dr["comp_pn"] + "')";
            
             qrySave = qrySave + "  exec sp_VendorSplitDuplicate '"+dr["SITE"]+"','"+dr["sku_code"]+"','"+dr["vendor_code"] +"','"+dr["split"]+"','"+ComboQtrCopy.SelectedText.ToString() + "','" + dr["comp_pn"] +"'";
	
            }

            qrySave = qrySave + " select @@error ";
            if (dtCopy.Rows.Count > 0)
            {
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlTransaction tran = cn.BeginTransaction("Transaction");
                SqlCommand cmd = new SqlCommand(qrySave, cn);
                cmd.Transaction = tran;

                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    tran.Commit();
                    MessageBox.Show("Updated successfully");
                    bindgrid();

                }
                catch (System.Exception exp)
                {
                    string error = exp.Message.ToString();
                    tran.Rollback();
                    MessageBox.Show("Failed to Update:" + error);
                }
                finally
                {
                    cn.Close();
                }
            }
            else
            {


            }

        }
        string currQtr;

        public void bindQrterCopy()
        {
            string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
            //string sql = " SELECT 'Q'+DATENAME(Quarter, GETDATE())+'-'+cast(YEAR(getdate()) as varchar) as Quarter, 1 as ord union SELECT 'Q'+DATENAME(Quarter,DATEADD(QQ,1, GETDATE()))+'-'+ cast(YEAR(DATEADD(QQ,1, GETDATE())) as varchar) as Quarter ,2 as ord  union SELECT 'Q'+DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))+'-'+ cast(YEAR(DATEADD(QQ,2, GETDATE())) as varchar) as Quarter,3 as ord  order by  ord ";
            string sql = "  SELECT * from m_app_vendorsplit_quarter  order by  ord  ";
            DataTable dtQuarter = getDataSet(sql, constr).Tables[0];
            currQtr = getDataSet(sql, constr).Tables[0].Rows[0][0].ToString();
            ComboQtrCopy.DataSource = dtQuarter;
            ComboQtrCopy.ValueMember = "Quarter";
            ComboQtrCopy.DisplayMember = "Quarter";
            ComboQtrCopy.DisplayLayout.Bands[0].Columns["ord"].Hidden = true;
            

            ComboQtr.DataSource = dtQuarter;
            ComboQtr.ValueMember = "Quarter";
            ComboQtr.DisplayMember = "Quarter";
            ComboQtr.DisplayLayout.Bands[0].Columns["ord"].Hidden = true;


           cmbQuarterCurSplit.DataSource = dtQuarter;
           cmbQuarterCurSplit.ValueMember = "Quarter";
           cmbQuarterCurSplit.DisplayMember = "Quarter";
           cmbQuarterCurSplit.DisplayLayout.Bands[0].Columns["ord"].Hidden = true;
   
        }

    
        private void cmbPartNum_RowSelected(object sender, RowSelectedEventArgs e)
        {
           
            
        }

        private void cmpCompPart_RowSelected(object sender, RowSelectedEventArgs e)
        {
        
        }

        private void cmbPartNum_Leave(object sender, EventArgs e)
        {
            if (cmbPartNum.Text != "" )
            {
                string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
                string strqry = " SELECT * FROM v_vendor_split_parts_list where sku_code='" + cmbPartNum.Text + "'";
                DataSet ds = getDataSet(strqry, constr);
                lblErrorMessage.Text = "";
                this.txtdesc.Text = "";
                this.lblPlanType.Text = "";
                bool validPart = false;
                if (ds != null)
                {
                    if (ds.Tables.Count > 0)
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            validPart = true;
                            lblErrorMessage.Text = ds.Tables[0].Rows[0]["ErrorMessage"].ToString();
                            this.txtdesc.Text = ds.Tables[0].Rows[0]["description"].ToString();
                            this.lblPlanType.Text = ds.Tables[0].Rows[0]["PlanType"].ToString();
                        }

                }
                if (!validPart)
                {
                    MessageBox.Show("Not Valid part", "Error Message", MessageBoxButtons.OK);
                    cmbPartNum.Text = "";
                    cmbPartNum.Focus();
                }
            }
        }

        private void grdQuarterCombo_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {

        }

        private void grdQuarterCombo_RowSelected(object sender, RowSelectedEventArgs e)
        {
            if (grdQuarterCombo.Text != "")
            {
                string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;

                string strqry = " select * from v_vendor_splt_display where   [Part Number]='" + cmbPartNum.Text + "' and  QUARTER>=(SELECT     cast(YEAR(getdate()) AS varchar) + 'Q' + replicate('0', (2 - len(DATENAME(Quarter, GETDATE())))) + DATENAME(Quarter, GETDATE()) )  ";
                if (grdQuarterCombo.Text != "ALL")

                    strqry = strqry + "  and     quarter='" + grdQuarterCombo.Text + "'";

                DataSet ds = getDataSet(strqry, constr);
                grdVendorSplit.DataSource = ds.Tables[0];
            }

        }

        private void btnFilterCurrSplit_Click(object sender, EventArgs e)
        {
            bindgridCurrSplit();
        }



        private void bindgridCurrSplit()
        {

            try
            {
               
                    string sqlGrid = "SELECT *   FROM [dbo].[m_app_vendorsplit_AllPart] where vendorCode is not null ";
                    string whereClause="";
                    if (cmbQuarterCurSplit.Text != "")
                        whereClause = whereClause + " and  Quarter='" + cmbQuarterCurSplit.Text + "'";
                    if (cmbSiteCurSplit.Text !="")
                        whereClause = whereClause + "   and site='" + cmbSiteCurSplit.Text + "' ";
                    if (cmbVendorCurSplit.Text != "")
                        whereClause = whereClause + "  and  vendorCode='" + cmbVendorCurSplit.Text + "'";

                    if (txtPrimGPN.Text != "")
                        whereClause = whereClause + "   and PrimGPN like '" + txtPrimGPN.Text + "%' ";

                    if (txtSplitGPN.Text != "")
                        whereClause = whereClause + "  and  SplitGPN like '" + txtSplitGPN.Text + "%'";

                    sqlGrid = sqlGrid + whereClause;
                    string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
                    dsGrid = getDataSet(sqlGrid, constr);
                    dtRepeater = getDataSet(sqlGrid, constr).Tables[0];

              

               
                grdVendorCurrentSplit.DataSource = dtRepeater;
                grdVendorCurrentSplit.DisplayLayout.Bands[0].Columns[3].Width = 400;
              
            }
            catch
            { }
        }

        private void btnDupAll_Click(object sender, EventArgs e)
        {
            dtCopy = repeaterVendorSplit.DataSource as DataTable;

            string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
            string qrySave = "";
            DataTable dtQt=ComboQtrCopy.DataSource as DataTable;
            foreach (DataRow dr in dtCopy.Rows)
            {
                qrySave = qrySave + " IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='AT' and sku_code='"+dr["sku_code"]+"' and vendor_code='"+dr["vendor_code"]+"' and quarter='"+dtQt.Rows[0][0].ToString()+"')  Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('AT','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[0][0].ToString() + "','" + dr["comp_pn"] + "')";
                qrySave = qrySave + "  IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='NL' and sku_code='" + dr["sku_code"] + "' and vendor_code='" + dr["vendor_code"] + "' and quarter='" + dtQt.Rows[0][0].ToString() + "')  Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('NL','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[0][0].ToString() + "','" + dr["comp_pn"] + "')";
          //      qrySave = qrySave + "  IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='SG' and sku_code='" + dr["sku_code"] + "' and vendor_code='" + dr["vendor_code"] + "' and quarter='" + dtQt.Rows[0][0].ToString() + "')  Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('SG','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[0][0].ToString() + "','" + dr["comp_pn"] + "')";

                qrySave = qrySave + "  IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='AT' and sku_code='" + dr["sku_code"] + "' and vendor_code='" + dr["vendor_code"] + "' and quarter='" + dtQt.Rows[1][0].ToString() + "')  Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('AT','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[1][0].ToString() + "','" + dr["comp_pn"] + "')";
                qrySave = qrySave + "  IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='NL' and sku_code='" + dr["sku_code"] + "' and vendor_code='" + dr["vendor_code"] + "' and quarter='" + dtQt.Rows[1][0].ToString() + "')  Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('NL','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[1][0].ToString() + "','" + dr["comp_pn"] + "')";
            //    qrySave = qrySave + " IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='SG' and sku_code='" + dr["sku_code"] + "' and vendor_code='" + dr["vendor_code"] + "' and quarter='" + dtQt.Rows[1][0].ToString() + "')   Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('SG','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[1][0].ToString() + "','" + dr["comp_pn"] + "')";

                qrySave = qrySave + "  IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='AT' and sku_code='" + dr["sku_code"] + "' and vendor_code='" + dr["vendor_code"] + "' and quarter='" + dtQt.Rows[2][0].ToString() + "')  Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('AT','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[2][0].ToString() + "','" + dr["comp_pn"] + "')";
                qrySave = qrySave + "  IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='NL' and sku_code='" + dr["sku_code"] + "' and vendor_code='" + dr["vendor_code"] + "' and quarter='" + dtQt.Rows[2][0].ToString() + "')  Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('NL','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[2][0].ToString() + "','" + dr["comp_pn"] + "')";
             //   qrySave = qrySave + "  IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='SG' and sku_code='" + dr["sku_code"] + "' and vendor_code='" + dr["vendor_code"] + "' and quarter='" + dtQt.Rows[2][0].ToString() + "')  Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('SG','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[2][0].ToString() + "','" + dr["comp_pn"] + "')";

                qrySave = qrySave + "  IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='AT' and sku_code='" + dr["sku_code"] + "' and vendor_code='" + dr["vendor_code"] + "' and quarter='" + dtQt.Rows[3][0].ToString() + "')  Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('AT','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[3][0].ToString() + "','" + dr["comp_pn"] + "')";
                qrySave = qrySave + "  IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='NL' and sku_code='" + dr["sku_code"] + "' and vendor_code='" + dr["vendor_code"] + "' and quarter='" + dtQt.Rows[3][0].ToString() + "')  Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('NL','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[3][0].ToString() + "','" + dr["comp_pn"] + "')";
             //   qrySave = qrySave + "  IF not EXISTS (SELECT 1 FROM [MIMDIST].[dbo].[m_vendor_split_table] where site='SG' and sku_code='" + dr["sku_code"] + "' and vendor_code='" + dr["vendor_code"] + "' and quarter='" + dtQt.Rows[3][0].ToString() + "')  Insert into [dbo].[m_vendor_split_table]([site],[sku_code],[vendor_code],[split],[quarter],comp_pn) values('SG','" + dr["sku_code"] + "','" + dr["vendor_code"] + "','" + dr["split"] + "','" + dtQt.Rows[3][0].ToString() + "','" + dr["comp_pn"] + "')";
            }

            qrySave = qrySave + " select @@error ";
            if (dtCopy.Rows.Count > 0)
            {
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlTransaction tran = cn.BeginTransaction("Transaction");
                SqlCommand cmd = new SqlCommand(qrySave, cn);
                cmd.Transaction = tran;
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    tran.Commit();
                    MessageBox.Show("Updated successfully");
                    bindgrid();

                }
                catch (System.Exception exp)
                {
                    string error = exp.Message.ToString();
                    tran.Rollback();
                    MessageBox.Show("Failed to Update:" + error);
                }
                finally
                {
                    cn.Close();
                }
            }
            else
            {


            }
        }

     
       

     
       
    }
}
