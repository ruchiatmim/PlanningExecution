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
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Drawing.Printing;
using System.Deployment.Application;
using System.Net;
using System.Net.Mail;
using Microsoft.Office.Interop.Outlook;

namespace Version3
{
    public partial class RepeaterTest : Form
    {
        public RepeaterTest()
        {
            InitializeComponent();          
        }
        string Exportfilename = "";
        private void releasesBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {            
            DateTime v_release_date = System.DateTime.Now;
            decimal v_quantity = 0;
            decimal v_received = 0;
            string v_status = "O";
            string v_confirm_date = "";
            string v_confirmed = "N";
            string v_ord_line = "";
            string v_po_line = "";
            string v_m_ship_method = "";
            DateTime v_m_factory_date = System.DateTime.Now;
            DateTime m_confirm_date = System.DateTime.Now;
            string v_m_ship_confirm = "";
            string v_m_tracking = "";

            string v_carrier = "";
            DateTime v_due_date;
            int row_id = 0;
            string v_location = "";
            string v_site = "";
            string vend_code = "";
            string quarter = "";
            string v_part_num = "";
            string ship_method_code = "";

            string insertqryIn = " insert into releases (po_no,part_no,location,part_type,release_date,quantity,received,status,confirm_date,confirmed,lb_tracking,conv_factor,prev_qty,po_key,due_date,ord_line,po_line,m_ship_method,m_factory_date,m_ship_confirm,m_tracking,carrier,m_confirmed_date,notes) ";
            string insertqry = "";
            string updateqry = "";
            bool save = true;
            string rel_notes = "";
            v_site = this.txtLocation.Text.ToString().Replace("-BUILD", "");
            v_part_num = this.txtPartNumber.Text.ToString();
            vend_code = this.textBoxVendor.Text.ToString();
            quarter = this.textBoxQuarter.Text.ToString();

            for (int i = 0; i < RepeaterPO.ItemCount; i++)
            {
                RepeaterPO.CurrentItemIndex = i;
                Microsoft.VisualBasic.PowerPacks.DataRepeaterItem item = RepeaterPO.CurrentItem;
                
                if (RepeaterPO.CurrentItem.Controls["textBoxRowID"].Text.ToString() != "")
                    row_id = Convert.ToInt32(RepeaterPO.CurrentItem.Controls["textBoxRowID"].Text.ToString());
                else
                    row_id = 0;

                v_release_date = v_release_date.AddHours(i + 1);
                
                CheckBox chkstatus = item.Controls["checkBoxStatus"] as CheckBox; //RepeaterPO.CurrentItem.Controls["checkBoxStatus"] as CheckBox;

                if (chkstatus.Checked)
                {
                    v_status = "C";
                }
                else
                {
                    v_status = "O";
                }


                if (RepeaterPO.CurrentItem.Controls["txtQty"].Text != "")
                {
                    v_quantity = Convert.ToDecimal((item.Controls["txtQty"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value);
                }


                if (RepeaterPO.CurrentItem.Controls["textBoxRecQty"].Text.Replace(",", "") != "")
                {
                    v_received = Convert.ToDecimal(item.Controls["textBoxRecQty"].Text);
                }


            //   int days = Convert.ToInt32(item.Controls["time_lead"].Text);
            //   v_m_factory_date = Convert.ToDateTime((item.Controls["dateTimePickerFacoryDate"] as DateTimePicker).Value);
                /*   if (RadFactDate.Checked)
                   {
                       v_due_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["DateTimePickerRelease_date"] as DateTimePicker).Value);
                       v_m_factory_date = v_due_date.AddDays(-1 * days);          
                   
                   }
                   else
                   {
                       v_m_factory_date = Convert.ToDateTime((item.Controls["dateTimePickerFacoryDate"] as DateTimePicker).Value);
                       v_due_date = v_m_factory_date.AddDays(days);
                                   
                   }  */


              //  if (RadFactDate.Checked)
               // {
               //   v_due_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
               //   v_m_factory_date = v_due_date.AddDays(-1 * days);
                //}
               // else
               // {
                   
                    v_m_factory_date = Convert.ToDateTime((item.Controls["ultraDateTimeFactoryDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                    v_due_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);

                //}


                rel_notes = (item.Controls["txtReleaseNote"] as Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor).Text;
                CheckBox chkConfirm = item.Controls["checkBoxConfirmed"] as CheckBox;
                if (chkConfirm.Checked)
                {
                    v_confirmed = "Y";
                    v_confirm_date = System.DateTime.Now.ToString();
                    if ((item.Controls["ultraDateTimeConfirmDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value != null)
                    {
                        m_confirm_date = Convert.ToDateTime((item.Controls["ultraDateTimeConfirmDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                    }
                    else
                    {
                        m_confirm_date = Convert.ToDateTime("01/01/1900");
                    }
                }
                else
                {
                    v_confirmed = "N";
                    m_confirm_date = Convert.ToDateTime("01/01/1900");
                }
              

                if ((chkConfirm.Checked) && (item.Controls["ultraDateTimeConfirmDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value == null)
                {
                    save = false;
                    goto Nosave;
                }

                v_m_ship_method = item.Controls["comboBoxShipMethod"].Text;
                ComboBox cmbShipMethod = item.Controls["comboBoxShipMethod"] as ComboBox;
                
                if (cmbShipMethod.Text.ToString() != "")
                {
                    ship_method_code = dtShipMeth.Select("ship_meth='" + cmbShipMethod.Text.ToString() + "'")[0][1].ToString();
                }

                // Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["DateTimePickerRelease_date"]as DateTimePicker).Value);
                CheckBox chkShipConfirm = item.Controls["checkBoxShipConf"] as CheckBox;

                if (chkShipConfirm.Checked)
                {
                    v_m_ship_confirm = "Y";
                }
                else
                {
                    v_m_ship_confirm = "N";
                }
                
                if (item.Controls["comboBoxCarrier"].Text.ToString() != "")
                {
                    v_carrier = dtCarrier.Select("ship_via_name='" + item.Controls["comboBoxCarrier"].Text.ToString() + "'")[0][0].ToString();
                }
              
                //  v_m_tracking = item.Controls["textBoxTracking"].Text;
                // v_due_date = getDueDate(v_site, vend_code, quarter, v_part_num, v_m_factory_date, v_m_ship_method);

                if (Convert.ToInt32(txtTrackTotQty.Value.ToString()) != 0)
                {
                    if (v_quantity != Convert.ToInt32(txtTrackTotQty.Value.ToString()))
                    {
                        save = false;
                        goto Nosave;
                    }
                }                                
                if (v_status != "C")
                    {
                        if (row_id == 0)
                        {
                            //insert into releases (po_no,part_no,location,part_type,release_date,quantity,received,status,confirm_date,confirmed,lb_tracking,conv_factor,prev_qty,po_key,due_date,ord_line,po_line,m_ship_method,m_factory_date,m_ship_confirm,m_tracking,carrier) ";
                            insertqry = insertqry + insertqryIn + " select po_no,part_no,location,type,'" + v_release_date + "','" + v_quantity + "','" + v_received + "','" + v_status + "','" + v_confirm_date + "','" + v_confirmed + "',lb_tracking,conv_factor,prev_qty,po_key,'" + v_due_date + "','" + v_ord_line + "',line,'" + ship_method_code + "','" + v_m_factory_date + "','" + v_m_ship_confirm + "','" + v_m_tracking + "','" + v_carrier + "','" + m_confirm_date + "','" + rel_notes + "' from pur_list  where po_no=" + txtPO.Text.Trim() + "  and line=" + txtLineNum.Text.Trim();
                        }
                        else
                        {
                            updateqry = updateqry + " update releases set  due_date='" + v_due_date +/* "',release_date='" + v_release_date +*/ "',quantity=" + v_quantity + ",received=" + v_received + ",status='" + v_status + "',confirm_date='" + v_confirm_date + "',confirmed='" + v_confirmed + "',m_ship_method='" + ship_method_code + "',m_factory_date='" + v_m_factory_date + "',m_ship_confirm='" + v_m_ship_confirm + "',m_tracking='" + v_m_tracking + "',carrier='" + v_carrier + "',m_confirmed_date='" + m_confirm_date+ "',notes='" + rel_notes + "' where row_id=" + row_id;
                        }
                    }                            
            }

            Nosave:
            string qry = insertqry + updateqry + " update purchase set m_revision=" + txtRev.Value.ToString() + "  where po_no = " + txtPO.Text.Trim();
            
            if (save)
            {
                string constr = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlTransaction tran = cn.BeginTransaction("Transaction");
                SqlCommand cmd = new SqlCommand(qry, cn);
                cmd.Transaction = tran;

                try
                {

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    tran.Commit();
                    MessageBox.Show("Updated successfully");

                }
                catch (System.Exception exp)
                {
                    string error = exp.Message.ToString();
                    tran.Rollback();
                    MessageBox.Show("Failed to Update:" + error);
                }
                cn.Close();
                bindgrid();
            }
            else
            {
                MessageBox.Show("Can not be updated. Confirm date is blank or quantity is not same as tracker quantity");
            }

        }


        public int getDueDate(string site, string vend_code, string quarter, string part_num,string fob, string ship_method_code)
        {
            int days = 0;
            string sqlquery = " SELECT  lead_time   FROM [MIMDIST].[dbo].[m_app_po_release_lead_time] where site='" + site + "' and FOB='"+fob+"' and ship_meth='"+ship_method_code+"'";// select trans_lt_at from m_app_vendor_portal_trans_lt where site='" + site + "' and vend_code='" + vend_code + "' and quote_qtr='" + quarter + "' and part_num='" + part_num + "' and ship_meth_at='" + ship_method_code + "'";
            string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlquery, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            if (ds.Tables.Count > 0)
                if (ds.Tables[0].Rows.Count > 0)
                {
                    days = Convert.ToInt32(ds.Tables[0].Rows[0]["lead_time"].ToString());
                }
            //  MessageBox.Show(days.ToString());
            //  DateTime release_date=  fact_date.AddDays(days);

            //return release_date;
            return days;

        }

        private void RepeaterTest_Load(object sender, EventArgs e)
        {           
            // TODO: This line of code loads data into the 'dataSet4.releases_tracking' table. You can move, or remove it, as needed.
            // this.releases_trackingTableAdapter.Fill(this.dataSet4.releases_tracking);
            // TODO: This line of code loads data into the 'dataSet4.releases' table. You can move, or remove it, as needed.
            // this.releasesTableAdapter.Fill(this.dataSet4.releases);
            // TODO: This line of code loads data into the 'dataSet1.arshipv' table. You can move, or remove it, as needed.
            // this.arshipvTableAdapter1.Fill(this.dataSet1.arshipv);
            // TODO: This line of code loads data into the 'dataSet1.Ship_meth' table. You can move, or remove it, as needed.
            //this.ship_methTableAdapter.Fill(this.dataSet1.Ship_meth);
            // arshipvTableAdapter.Fill(this.dataSet1.arshipv);         
        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            repeaterTracker.DataSource = null;
            try
            {
                if (txtLineStatus.Text.Trim() == "OPEN")
                {
                    //releasesBindingSource.AddNew();
                    RepeaterPO.AddNew();
                    (RepeaterPO.CurrentItem.Controls["checkBoxStatus"] as CheckBox).Checked = false;
                    (RepeaterPO.CurrentItem.Controls["txtQty"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).ReadOnly = false;
                    //  (item.Controls["textBoxTracking"] as TextBox).ReadOnly = true;
                    //   (item.Controls["comboBoxCarrier"] as ComboBox).Enabled = false;
                    (RepeaterPO.CurrentItem.Controls["comboBoxShipMethod"] as ComboBox).Enabled = true;
                    (RepeaterPO.CurrentItem.Controls["comboBoxShipMethod"] as ComboBox).SelectedIndex = 0;
                    (RepeaterPO.CurrentItem.Controls["checkBoxConfirmed"] as CheckBox).Enabled = true;
                    (RepeaterPO.CurrentItem.Controls["checkBoxShipConf"] as CheckBox).Enabled = true;
                    (RepeaterPO.CurrentItem.Controls["checkBoxStatus"] as CheckBox).Enabled = true;
                    (RepeaterPO.CurrentItem.Controls["checkBoxStatus"] as CheckBox).Checked= false;
                    (RepeaterPO.CurrentItem.Controls["checkBoxConfirmed"] as CheckBox).Checked = false;
                    (RepeaterPO.CurrentItem.Controls["checkBoxShipConf"] as CheckBox).Checked = false;
                }
                else
                    MessageBox.Show("New release can not be added because purchase line is closed");
            }
            catch
            {

            }
        }

        private void copyToolStripButton_Click(object sender, EventArgs e)
        {

        }

        private void ultraButton1_Click(object sender, EventArgs e)
        {
            int foundIndex;
            string searchString;
            //searchString = txtVendor.Text;
            // Search for the string in the CustomerID field.
            //   ShipToBindingSource1.Filter = "po_no=" + searchString;
            //     foundIndex = ShipToBindingSource1.Find("po_no", searchString);
            //if (foundIndex > -1)
            //{
            //  //  dataRepeater1.CurrentItemIndex = foundIndex;
            //}
            //else
            //{
            //    MessageBox.Show("Item " + searchString + " not found.");
            //}

        }


        DataTable dtRelease;
        DateTime? nullDattime;
        int GetQuarterName(DateTime myDate)
        {
            return (myDate.Month - 1) / 3 + 1;
        }
        string internal_notes = "";
            string external_notes;
        public void bindgrid()
        {

            loading = true;
            clearALL();
            repeaterTracker.DataSource = null;

            //string sql = " SELECT  dbo.releases.row_id, dbo.purchase.vendor_no,  dbo.releases.po_no, dbo.releases.part_no, dbo.releases.location, case when dbo.pur_list.status='O' then 'OPEN' when dbo.pur_list.status='C' then 'CLOSED' else dbo.pur_list.status end  as line_status, dbo.releases.part_type, dbo.releases.release_date, dbo.releases.quantity, dbo.releases.received, case when dbo.releases.status ='C' then 'true' else 'false' end as status, dbo.releases.confirm_date, case when dbo.releases.confirmed='Y' then 'true' else 'false' end as confirmed , dbo.releases.lb_tracking, dbo.releases.conv_factor, dbo.releases.prev_qty, dbo.releases.po_key, dbo.releases.row_id, dbo.releases.due_date, dbo.releases.ord_line, dbo.releases.po_line, dbo.releases.receipt_batch_no, dbo.releases.m_ship_method, dbo.releases.m_confirmed_date, dbo.releases.m_factory_date, case when dbo.releases.m_ship_confirm='Y' then 'true' else 'false' end as m_ship_confirm, dbo.releases.m_tracking, isnull(dbo.releases.carrier,purchase.ship_via) as carrier, dbo.pur_list.description,dbo.purchase.user_def_fld2 as quarter,datediff(day,dbo.releases.m_factory_date,dbo.releases.due_date) as diff, dbo.purchase.m_revision as revision   FROM         dbo.releases INNER JOIN  dbo.pur_list ON dbo.releases.po_no = dbo.pur_list.po_no AND dbo.releases.po_line = dbo.pur_list.line INNER JOIN dbo.purchase ON dbo.pur_list.po_no = dbo.purchase.po_no WHERE     (dbo.purchase.status = 'O') AND (dbo.purchase.m_po_type = 3) AND (dbo.purchase.void = 'N') AND (dbo.purchase.date_of_order > GETDATE() - 180)  ";            //    sql = sql + " and dbo.releases.part_no='" + txtPartNum.Text.Trim() + "'";
            string sql = " SELECT  * from dbo.m_app_po_release    ";

            //if  ( this.txtVendor.Text.Trim() != "")
            //    sql = sql + " and vendor_no='" + this.txtVendor.Text.Trim() + "'";
            this.RepeaterPO.DataSource = null;
            if (this.txtPO.Text.Trim() == "" || this.txtLineNum.Text.Trim() == "")
            {
                MessageBox.Show("Enter PO and Line#");
            }
            else
            {
                if (this.txtPO.Text.Trim() != "")
                    sql = sql + " where po_no='" + txtPO.Text.Trim() + "'";
                if (this.txtLineNum.Text.Trim() != "")
                    sql = sql + "  and  po_line='" + txtLineNum.Text.Trim() + "'";
                if (chkRelStatus.Checked)
                    sql = sql + "  and  status='false'";
                else
                    sql = sql + "  and  status='true'";
                sql = sql + " ORDER BY  due_date Asc   SELECT sum([quantity]) as tot_quantity, sum([received]) as tot_received   FROM [dbo].[m_app_po_release]  ";

                if (this.txtPO.Text.Trim() != "")
                    sql = sql + " where po_no='" + txtPO.Text.Trim() + "'";
                if (this.txtLineNum.Text.Trim() != "")
                    sql = sql + " and po_line='" + txtLineNum.Text.Trim() + "'";


                string constr = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;

                DataSet dsRelease = getDataSet(sql, constr);
                dtRelease = dsRelease.Tables[0];//.Select("po_line=1").CopyToDataTable();
                
                // releasesBindingSource.DataSource = dtRelease;
                //comboBox1.DataBindings.Add("SelectedItem", releasesBindingSource, "OverflowBehaviour", false, DataSourceUpdateMode.OnValidation);
                
                if (dtRelease != null)
                {
                    if (dtRelease.Rows.Count > 0)
                    {
                        if (dtRelease.Rows[0]["line_status"].ToString() == "CLOSED")
                        {
                            MessageBox.Show("Line is Closed. It must be opened");
                        }
                        else
                        {
                            txtPartNumber.Text = dtRelease.Rows[0]["part_no"].ToString();
                            this.txtLocation.Text = dtRelease.Rows[0]["location"].ToString();
                            this.textBoxDescription.Text = dtRelease.Rows[0]["description"].ToString();
                            this.txtLineStatus.Text = dtRelease.Rows[0]["line_status"].ToString();
                            this.textBoxQuarter.Text = dtRelease.Rows[0]["quarter"].ToString();
                            this.textBoxVendor.Text = dtRelease.Rows[0]["vendor_no"].ToString();
                            this.txtRev.Value = Convert.ToInt32(dtRelease.Rows[0]["revision"].ToString());
                            this.textBoxMfgLT.Text = dtRelease.Rows[0]["mfg_lt"].ToString();
                            this.textBoxOrderMult.Text = dtRelease.Rows[0]["ord_mult"].ToString();
                            this.textBoxQuotePrice.Text = dtRelease.Rows[0]["m_quote_price"].ToString();
                            this.textBoxFOB.Text = dtRelease.Rows[0]["FOB"].ToString();
                            this.textBoxShipMeth.Text = dtRelease.Rows[0]["ship_meth"].ToString();
                            textBoxCarrier.Text = dtRelease.Rows[0]["default_carrier"].ToString();
                            internal_notes = dtRelease.Rows[0]["Internal_Notes"].ToString().Trim();
                            external_notes = dtRelease.Rows[0]["External_Notes"].ToString().Trim();
                            if (internal_notes.Trim() != "")
                                pictureBox1.Load("image\\notepad_filled.png");
                            else
                                pictureBox1.Load("image\\notepad_blank.png");
                            if (external_notes.Trim() != "")
                                pictureBox2.Load("image\\notepad_filled.png");
                            else
                                pictureBox2.Load("image\\notepad_blank.png");

                            textBoxTotRecQty.Value = dsRelease.Tables[1].Rows[0]["tot_received"].ToString();
                            textBoxTotOrdQty.Value = dsRelease.Tables[1].Rows[0]["tot_quantity"].ToString();
                            if (dtRelease.Rows[0]["E2OpenFlag"].ToString().Trim() == "Y")
                            {
                              //  BtnE2Open.Appearance.BackColor = Color.Green;
                                BtnE2Open.Appearance.ForeColor = Color.RoyalBlue;
                            }
                            else
                            {
                                BtnE2Open.Appearance.ForeColor = Color.Black;
                               // BtnE2Open.Appearance.ForeColor = Color.White;
                            }
                          //  releasesBindingSource.DataSource = dtRelease;
                            //comboBox1.DataBindings.Add("SelectedItem", rele
                            try
                            {
                                this.RepeaterPO.DataSource = dtRelease;                               
                            }
                            catch (System.Exception exp)
                            {
                                //MessageBox.Show(exp.ToString());
                            }
                            finally
                            {
                                this.RepeaterPO.CurrentItemIndex = 0;
                                loadTracker();
                            }

                        }
                    }
                }
            }
            loading = false;
        }

        public void clearALL()
        {
            dtRelease = null;
            txtPartNumber.Text = "";
            this.txtLocation.Text = "";
            this.textBoxDescription.Text = "";
            this.txtLineStatus.Text = "";
            this.textBoxQuarter.Text = "";
            this.textBoxVendor.Text = "";
            this.txtRev.Value = null;
            this.RepeaterPO.DataSource = null;

        }
        bool loadLeadTime = false;
        bool loading = false;
        private void btnFilter_Click(object sender, EventArgs e)
        {
          
            bindgrid();
            txtTrackTotQty.Value = 0;    
            
            //string sql = "SELECT     dbo.releases.po_no, dbo.releases.part_no, dbo.releases.location, dbo.releases.part_type, dbo.releases.release_date, dbo.releases.quantity, dbo.releases.received, dbo.releases.status, dbo.releases.confirm_date, dbo.releases.confirmed, dbo.releases.lb_tracking, dbo.releases.conv_factor, dbo.releases.prev_qty, dbo.releases.po_key, dbo.releases.row_id, dbo.releases.due_date, dbo.releases.ord_line, dbo.releases.po_line, dbo.releases.receipt_batch_no, dbo.releases.m_ship_method, dbo.releases.m_confirmed_date, dbo.releases.m_factory_date, dbo.releases.m_ship_confirm, dbo.releases.m_tracking, dbo.releases.carrier, dbo.purchase.status AS Expr1, dbo.purchase.m_po_type, dbo.purchase.void, dbo.purchase.date_of_order FROM         dbo.releases INNER JOIN dbo.purchase ON dbo.releases.po_no = dbo.purchase.po_no WHERE     (dbo.purchase.status = 'O') AND (dbo.releases.status = 'O') AND (dbo.purchase.m_po_type = 3) AND (dbo.purchase.void = 'N') AND (dbo.purchase.date_of_order > GETDATE() - 180) ";
           
           
        }

        public void btn_click(object sender, EventArgs e)
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
        

        public void repeater_DrawItem(object sender, Microsoft.VisualBasic.PowerPacks.DataRepeaterItemEventArgs e)
        {

            
        }

        private void dataRepeater1_CurrentItemIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataRepeater1_DrawItem(object sender, Microsoft.VisualBasic.PowerPacks.DataRepeaterItemEventArgs e)
        {


        }

        private void fillByToolStripButton_Click(object sender, EventArgs e)
        {
           

        }

        private void BindingSourceShipMeth_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void CarrierBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {


            var combo = (ComboBox)sender;
            var DataRepeaterItem = (Microsoft.VisualBasic.PowerPacks.DataRepeaterItem)combo.Parent;

            //Update dataset
            try
            {
                if (dtRelease.Rows[DataRepeaterItem.ItemIndex]["carrier"].ToString() != combo.SelectedItem.ToString())
                {
                    dtRelease.Rows[DataRepeaterItem.ItemIndex]["carrier"] = dtCarrier.Select("ship_via_name='" + combo.Text.ToString() + "'")[0][0].ToString();
                }
            }
            catch
            {
            }


            ComboBox comboCarrier =this.RepeaterPO.CurrentItem.Controls["comboBoxCarrier"] as ComboBox;
            if (comboCarrier.SelectedValue != null && textBoxCarrier.Text.ToString() != "")
            {
                if (comboCarrier.SelectedValue.ToString() != textBoxCarrier.Text.ToString())
                {
                    RepeaterPO.CurrentItem.Controls["textBoxWarning"].Text = "*Default Carrier exception";
                }
                else
                {
                    RepeaterPO.CurrentItem.Controls["textBoxWarning"].Text = "";
                }
            }

            // DateTime fact_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["dateTimePickerFacoryDate"] as DateTimePicker).Value);
            //= Convert.ToDateTime(fact_date_string);
            //(RepeaterPO.CurrentItem.Controls["DateTimePickerRelease_date"] as DateTimePicker).Value = getDueDate(v_site, vend_code, quarter, v_part_num, fact_date, ship_method_code);
            //(RepeaterPO.CurrentItem.Controls["DateTimePickerRelease_date"] as DateTimePicker).Text = getDueDate(v_site, vend_code, quarter, v_part_num, fact_date, ship_method_code).ToString();
        }

        private void textBoxDescription_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void BindingSourceShipMeth_CurrentChanged_1(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            if (dtRelease.Rows.Count > 0)
            {
                DeleteLine();
            }
            try
            {
                RepeaterPO.RemoveAt(RepeaterPO.CurrentItemIndex);
            }
            catch { }
            bindgrid();
         }

        public void DeleteLine()
        {
            if (RepeaterPO.CurrentItem.Controls["textBoxRowID"].Text.Trim() != "")
            {
                string sql = "delete from releases where row_id=" + RepeaterPO.CurrentItem.Controls["textBoxRowID"].Text + " select @@error ";
                string constr = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;
                DataSet ds = getDataSet(sql, constr);
                if (ds.Tables[0].Rows[0][0].ToString() == "0")
                {
                    MessageBox.Show("Item Deleted");
                  
                }
                else
                {
                    MessageBox.Show("Error to delete Item");
                }
            }
        }



        private void comboBoxShipMethod_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (RepeaterPO.CurrentItemIndex == 0)
            {
                MessageBox.Show("ind");
            }

        }

        private void comboBoxShipMethod_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            string v_site = this.txtLocation.Text.ToString().Replace("-BUILD", "");
            string v_part_num = this.txtPartNumber.Text.ToString();
            string vend_code = this.textBoxVendor.Text.ToString();
            string quarter = this.textBoxQuarter.Text.ToString();
            string ship_method_code = RepeaterPO.CurrentItem.Controls["comboBoxShipMethod"].Text;
            string fob = this.textBoxFOB.Text;
            int days = getDueDate(v_site, vend_code, quarter, v_part_num,fob, ship_method_code);

            (RepeaterPO.CurrentItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value = days;
            var combo = (ComboBox)sender;
            var DataRepeaterItem = (Microsoft.VisualBasic.PowerPacks.DataRepeaterItem)combo.Parent;

            //Update dataset
            try
            {
                if (dtRelease.Rows[DataRepeaterItem.ItemIndex]["m_ship_method"].ToString() != combo.SelectedItem.ToString())
                {
                    dtRelease.Rows[DataRepeaterItem.ItemIndex]["m_ship_method"] = Convert.ToInt32(dtShipMeth.Select("ship_meth='" + combo.Text.ToString() + "'")[0][1].ToString());
                }
            }
            catch 
            { 
            
            }

            //= Convert.ToDateTime(fact_date_string);

            if (RadFactDate.Checked)
            {

                DateTime fact_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                
                //if (Convert.ToInt32((RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value) != Convert.ToInt32((RepeaterPO.CurrentItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value))
                //    (RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Appearance.BackColor = System.Drawing.Color.Orange;
                //else
                //    (RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Appearance.BackColor = System.Drawing.Color.White;

                (RepeaterPO.CurrentItem.Controls["ultraDateTimeFactoryDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value = fact_date.AddDays(-1 * days);
                (RepeaterPO.CurrentItem.Controls["ultraDateTimeFactoryDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Text = fact_date.AddDays(-1 * days).ToString();
            }
            else
            {

                CheckBox chk = RepeaterPO.CurrentItem.Controls["checkBoxConfirmed"] as CheckBox;
                if (!chk.Checked)
                {
                    DateTime fact_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeFactoryDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                    
                    //if (Convert.ToInt32((RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value) != Convert.ToInt32((RepeaterPO.CurrentItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value))
                    //    (RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Appearance.BackColor = System.Drawing.Color.Orange;
                    //else
                    //    (RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Appearance.BackColor = System.Drawing.Color.White;
                    if (!this.loading)
                    {
                        (RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value = fact_date.AddDays(days);
                        (RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Text = fact_date.AddDays(days).ToString();
                    } 
                 
                }
                else
                {
                    DateTime fact_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeConfirmDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                    
                    //if (Convert.ToInt32((RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value) != Convert.ToInt32((RepeaterPO.CurrentItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value))
                    //    (RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Appearance.BackColor = System.Drawing.Color.Orange;
                    //else
                    //    (RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Appearance.BackColor = System.Drawing.Color.White;

                    int days1 = Convert.ToInt32((RepeaterPO.CurrentItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value);                    
                    //  (RepeaterPO.CurrentItem.Controls["time_lead"] as TextBox).Text = days1.ToString();                    
                    if (!this.loading)
                    {
                        (RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value = fact_date.AddDays(days1);
                        (RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Text = fact_date.AddDays(days1).ToString();
                    }

                }
            }

            getTimeLead();
        }

        DataTable dtShipMeth;
        DataTable dtCarrier;
        private void RepeaterPO_ItemCloned(object sender, Microsoft.VisualBasic.PowerPacks.DataRepeaterItemEventArgs e)
        {
            string constr = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;           
            try
            {
                ComboBox Cmb = e.DataRepeaterItem.Controls["comboBoxShipMethod"] as ComboBox;
                try
                {
                    DataSet ds = getDataSet("SELECT distinct ship_meth, ep_po_met_code AS code FROM  m_pe_transit_lt_master ", constr);
                    dtShipMeth = ds.Tables[0];
                    DataRow dr = dtShipMeth.NewRow();
                    dr[0] = "SELECT";
                    dr[1] = 0;
                    dtShipMeth.Rows.InsertAt(dr, 0);
                    if (Cmb.DataSource == null)
                    {
                        Cmb.DataSource = dtShipMeth;
                        Cmb.DisplayMember = "ship_meth";
                        Cmb.ValueMember = "ep_po_met_code";
                    }
                    Cmb.SelectedIndexChanged += new EventHandler(comboBoxShipMethod_SelectedIndexChanged_1);
                }
                catch
                {
                    Cmb.SelectedIndex = 0;
                }
            }
            catch { }

        //    int days = getDueDate(v_site, vend_code, quarter, v_part_num, ship_method_code);

            //  string ship_meth = dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["m_ship_method"].ToString();
            //   Cmb.SelectedValue = "2";// dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["ep_po_met_code"];
            //Cmb.SelectedValue = dtRelease..Columns["m_ship_method"];

            // Cmb.DataBindings.Add("SelectedValue", dtRelease, "m_ship_Method");
            string constrep = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;
            try
            {
                           ComboBox CmbCarrier = e.DataRepeaterItem.Controls["comboBoxCarrier"] as ComboBox;
                           dtCarrier = getDataSet("SELECT ship_via_code, ship_via_name,addr2 FROM arshipv", constrep).Tables[0];
                               if (CmbCarrier.DataSource == null)
                               {
                                   CmbCarrier.DisplayMember = "ship_via_name";
                                   CmbCarrier.ValueMember = "ship_via_code";
                                   CmbCarrier.DataSource = dtCarrier;
                               }
                           CmbCarrier.SelectedIndexChanged += new EventHandler(comboBox2_SelectedIndexChanged);
               
                /*       DateTimePicker dt = e.DataRepeaterItem.Controls["dateTimePickerFacoryDate"] as DateTimePicker;
                         dt.ValueChanged += new EventHandler(dateTimePickerFacoryDate_ValueChanged_1);

                         DateTimePicker dateRelease = e.DataRepeaterItem.Controls["DateTimePickerRelease_date"] as DateTimePicker;
                         dateRelease.ValueChanged += new EventHandler(DateTimePickerRelease_date_ValueChanged);              */

                Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtUltra = e.DataRepeaterItem.Controls["ultraDateTimeFactoryDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor;
                dtUltra.ValueChanged += new EventHandler(ultraDateTimeFactoryDate_ValueChanged);

                Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtUltraMRP = e.DataRepeaterItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor;
                dtUltraMRP.ValueChanged += new EventHandler(ultraDateTimeMRPDate_ValueChanged);

                CheckBox chkConfirm = e.DataRepeaterItem.Controls["checkBoxConfirmed"] as CheckBox;
                chkConfirm.CheckedChanged += new EventHandler(checkBoxConfirmed_CheckedChanged);


                Infragistics.Win.UltraWinEditors.UltraDateTimeEditor Confirm_date = e.DataRepeaterItem.Controls["ultraDateTimeConfirmDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor;
                Confirm_date.ValueChanged += new EventHandler(ultraDateTimeConfirmDate_ValueChanged);

                Infragistics.Win.UltraWinEditors.UltraNumericEditor txtorderQty = e.DataRepeaterItem.Controls["txtQty"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor;
                txtorderQty.TextChanged += new EventHandler(txtQty_ValueChanged);


                 //Infragistics.Win.UltraWinEditors.UltraNumericEditor txtleadQty = e.DataRepeaterItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor;
                 //txtleadQty.ValueChanged += new EventHandler(time_lead_ValueChanged);
                
                //ComboBox Cmb = e.DataRepeaterItem.Controls["comboBoxShipMethod"] as ComboBox;             

            }
            catch (System.Exception ex)
            {
                //   MessageBox.Show("error" + ex.Message.ToString());
            }

            try
            {
                CheckBox chkConfirm = e.DataRepeaterItem.Controls["checkBoxConfirmed"] as CheckBox;
                if (chkConfirm.Checked)
                {
                    e.DataRepeaterItem.Controls["ultraDateTimeConfirmDate"].Visible = true;
                }
                else
                {
                    e.DataRepeaterItem.Controls["ultraDateTimeConfirmDate"].Visible = false;
                }
            }
            catch { }
            //try
            //{
            //    ComboBox CmbCarrier = e.DataRepeaterItem.Controls["comboBoxCarrier"] as ComboBox;

            //    dtCarrier = getDataSet("SELECT ship_via_code, ship_via_name FROM arshipv", constr).Tables[0];
            //    if (CmbCarrier.DataSource == null)
            //    {
            //        CmbCarrier.DisplayMember = "ship_via_name";
            //        CmbCarrier.ValueMember = "ship_via_code";
            //        CmbCarrier.DataSource = dtCarrier;
            //    }
            //    CmbCarrier.SelectedIndexChanged += new EventHandler(comboBox2_SelectedIndexChanged);

            //    /*       DateTimePicker dt = e.DataRepeaterItem.Controls["dateTimePickerFacoryDate"] as DateTimePicker;
            //             dt.ValueChanged += new EventHandler(dateTimePickerFacoryDate_ValueChanged_1);

            //             DateTimePicker dateRelease = e.DataRepeaterItem.Controls["DateTimePickerRelease_date"] as DateTimePicker;
            //             dateRelease.ValueChanged += new EventHandler(DateTimePickerRelease_date_ValueChanged);              */

            //    Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtUltra = e.DataRepeaterItem.Controls["ultraDateTimeFactoryDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor;
            //    dtUltra.ValueChanged += new EventHandler(ultraDateTimeFactoryDate_ValueChanged);

            //    Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtUltraMRP = e.DataRepeaterItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor;
            //    dtUltraMRP.ValueChanged += new EventHandler(ultraDateTimeMRPDate_ValueChanged);

            //    //ComboBox Cmb = e.DataRepeaterItem.Controls["comboBoxShipMethod"] as ComboBox;             

            //}
            //catch (System.Exception ex)
            //{
            //    MessageBox.Show("error" + ex.Message.ToString());
            //}  


         
        }

        private void dateTimePickerFacoryDate_ValueChanged_1(object sender, EventArgs e)
        {

            try
            {
                string v_site = this.txtLocation.Text.ToString().Replace("-BUILD", "");
                string v_part_num = this.txtPartNumber.Text.ToString();
                string vend_code = this.textBoxVendor.Text.ToString();
                string quarter = this.textBoxQuarter.Text.ToString();
                string ship_method_code = RepeaterPO.CurrentItem.Controls["comboBoxShipMethod"].Text;
                DateTime fact_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["dateTimePickerFacoryDate"] as DateTimePicker).Text);
                string fob = this.textBoxFOB.Text;
                int days = getDueDate(v_site, vend_code, quarter, v_part_num,fob, ship_method_code);
                if (RadMRPDate.Checked)
                {
                    // (RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value = days;
                    (RepeaterPO.CurrentItem.Controls["DateTimePickerRelease_date"] as DateTimePicker).Value = fact_date.AddDays(days);
                    //  (RepeaterPO.CurrentItem.Controls["DateTimePickerRelease_date"] as DateTimePicker).Text = fact_date.AddDays(days).ToString();
                }
            }
            catch (System.Exception exe)
            {
            }

        }

        private void RepeaterPO_DrawItem(object sender, Microsoft.VisualBasic.PowerPacks.DataRepeaterItemEventArgs e)
        {

            try
            {
                ComboBox Cmb = e.DataRepeaterItem.Controls["comboBoxShipMethod"] as ComboBox;
                try
                {
                    if (dtRelease != null)
                    {
                        for (int i = 0; i < dtShipMeth.Rows.Count; i++)
                        {
                            // if (dtRelease.Rows.Count > e.DataRepeaterItem.ItemIndex)
                            if (dtShipMeth.Rows[i][1].ToString() == dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["m_ship_method"].ToString())
                            {
                                Cmb.SelectedIndex = i;
                                goto abc;
                            }
                        }
                    }
                }
                catch
                {
                    // MessageBox.Show("The number {0} was not found.Catch 1");
                    Cmb.SelectedIndex = 0;


                }
            }
            catch
            {
                // MessageBox.Show("The number {0} was not found.Catch 2");
                //Cmb.SelectedIndex = 0;
            }
              abc:  Console.WriteLine("The number {0} was not found." );
                try
                {

                    ComboBox CmbCarrier = e.DataRepeaterItem.Controls["comboBoxCarrier"] as ComboBox;
                     if (dtRelease != null)
                         for (int i = 0; i < this.dtCarrier.Rows.Count; i++)
                         {
                             if (dtRelease.Rows.Count > e.DataRepeaterItem.ItemIndex)
                                 if (dtCarrier.Rows[i][0].ToString() == dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["carrier"].ToString())
                                     CmbCarrier.SelectedIndex = i;

                         }
            }

                catch
                {
                    // MessageBox.Show("The number {0} was not found.Catch 2");
                    //Cmb.SelectedIndex = 0;
                }
            try
            {
                CheckBox chkConfirm = e.DataRepeaterItem.Controls["checkBoxConfirmed"] as CheckBox;
                try
                {
                    if (chkConfirm.Checked)
                    {
                        e.DataRepeaterItem.Controls["ultraDateTimeConfirmDate"].Visible = true;
                    }
                    else
                    {
                        e.DataRepeaterItem.Controls["ultraDateTimeConfirmDate"].Visible = false;
                    }
                }
                catch
                {
                    chkConfirm.Checked = false;
                }
               
               
            }
            catch {
               // MessageBox.Show("confirmed catch 2");
            }
            //  }
            try
            {
                string v_site = this.txtLocation.Text.ToString().Replace("-BUILD", "");
                string v_part_num = this.txtPartNumber.Text.ToString();
                string vend_code = this.textBoxVendor.Text.ToString();
                string quarter = this.textBoxQuarter.Text.ToString();
                string ship_method_code = e.DataRepeaterItem.Controls["comboBoxShipMethod"].Text;
                //   DateTime fact_date = Convert.ToDateTime((e.DataRepeaterItem.Controls["dateTimePickerFacoryDate"] as DateTimePicker).Text);
                string fob = this.textBoxFOB.Text;
                int days = getDueDate(v_site, vend_code, quarter, v_part_num,fob, ship_method_code);

                (e.DataRepeaterItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value = days;
         /*       if (Convert.ToInt32(dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["diff"].ToString()) != Convert.ToInt32((e.DataRepeaterItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value))
                {
                  //  (e.DataRepeaterItem.Controls["time_lead"] as TextBox).Text = dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["diff"].ToString();
                    (e.DataRepeaterItem.Controls["time_lead"] as TextBox).BackColor = System.Drawing.Color.Red;
                }
                else
                {
                  //  (e.DataRepeaterItem.Controls["time_lead"] as TextBox).Text = dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["diff"].ToString();
                    (e.DataRepeaterItem.Controls["time_lead"] as TextBox).BackColor = System.Drawing.Color.White;
                }
                */
                (e.DataRepeaterItem.Controls["ultraDateTimeMRPDate"] as  Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value = dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["due_date"].ToString();
            }
            catch(System.Exception exp) { }


           try
            {
                DateTime request_date = System.DateTime.Now;
               
               DateTime   mrp_date_value = Convert.ToDateTime((e.DataRepeaterItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                if ((e.DataRepeaterItem.Controls["checkBoxConfirmed"] as CheckBox).Checked)
                
                    request_date = Convert.ToDateTime((e.DataRepeaterItem.Controls["ultraDateTimeConfirmDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                else
                {
                    request_date = Convert.ToDateTime((e.DataRepeaterItem.Controls["ultraDateTimeFactoryDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                }
               

                TimeSpan ts = Convert.ToDateTime(mrp_date_value) - Convert.ToDateTime(request_date);
                if (ts.TotalDays == 0)
                    e.DataRepeaterItem.Controls["time_lead"].Text = "0";
                else
                {
                    e.DataRepeaterItem.Controls["time_lead"].Text = Convert.ToInt32(ts.TotalDays).ToString();
                }
                if (Convert.ToInt32(e.DataRepeaterItem.Controls["time_lead"].Text) != Convert.ToInt32((e.DataRepeaterItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value))
                {
                   //   (e.DataRepeaterItem.Controls["time_lead"] as TextBox).Text = dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["diff"].ToString();
                    e.DataRepeaterItem.Controls["time_lead"].ForeColor = System.Drawing.Color.Red;
                }
                else
                {
                   //   (e.DataRepeaterItem.Controls["time_lead"] as TextBox).Text = dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["diff"].ToString();
                    e.DataRepeaterItem.Controls["time_lead"].ForeColor = System.Drawing.Color.Black;
                }
            }
            catch { }
        }

        private void releasesBindingNavigator_RefreshItems(object sender, EventArgs e)
        {

        }

        private void comboBoxShipMethod_Leave(object sender, EventArgs e)
        {

        }

        private void checkBoxStatus_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if ((sender as CheckBox).Checked)
                {
                    (sender as CheckBox).Enabled = false;
                    Microsoft.VisualBasic.PowerPacks.DataRepeaterItem item = (sender as CheckBox).Parent as Microsoft.VisualBasic.PowerPacks.DataRepeaterItem;
                    // item.Controls["DateTimePickerRelease_date"].Enabled=false;
                    //   item.Controls["dateTimePickerFacoryDate"].Enabled = false;
                    (item.Controls["txtQty"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).ReadOnly = true;
                    //  (item.Controls["textBoxTracking"] as TextBox).ReadOnly = true;

                    (item.Controls["comboBoxShipMethod"] as ComboBox).Enabled = false;
                    (item.Controls["checkBoxConfirmed"] as CheckBox).Enabled = false;
                    (item.Controls["checkBoxShipConf"] as CheckBox).Enabled = false;
                    (item.Controls["checkBoxStatus"] as CheckBox).Enabled = false;
                    //(RepeaterPO.CurrentItem.Controls["dateTimePickerFacoryDate"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).ReadOnly=true;
                    // (RepeaterPO.CurrentItem.Controls["txtQty"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).ReadOnly=true;
                    //(RepeaterPO.CurrentItem.Controls["txtQty"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).ReadOnly=true;

                    //   int row_id = Convert.ToInt32((item.Controls["textBoxRowID"] as TextBox).Text);
                    //  closeRelease(row_id);
                }
                else
                {
                    (sender as CheckBox).Enabled = true;
                    Microsoft.VisualBasic.PowerPacks.DataRepeaterItem item = (sender as CheckBox).Parent as Microsoft.VisualBasic.PowerPacks.DataRepeaterItem;
                    // item.Controls["DateTimePickerRelease_date"].Enabled=false;
                    //   item.Controls["dateTimePickerFacoryDate"].Enabled = false;
                    (item.Controls["txtQty"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).ReadOnly = false;
                    //  (item.Controls["textBoxTracking"] as TextBox).ReadOnly = true;
                    //   (item.Controls["comboBoxCarrier"] as ComboBox).Enabled = false;
                    (item.Controls["comboBoxShipMethod"] as ComboBox).Enabled = true;
                    (item.Controls["checkBoxConfirmed"] as CheckBox).Enabled = true;
                    (item.Controls["checkBoxShipConf"] as CheckBox).Enabled = true;
                    (item.Controls["checkBoxStatus"] as CheckBox).Enabled = true;
                }
            }
            catch (System.Exception ex)
            { 
            
            }
        }

        public void closeRelease(int row_id)
        {
               string updateqry =" update releases set  status='C' where row_id=" + row_id;
                       
        
                   string constr = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;

                SqlConnection cn = new SqlConnection(constr);
                cn.Open();

                SqlCommand cmd = new SqlCommand(updateqry, cn);
                try
                {

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                 

                    MessageBox.Show("Updated successfully");

                }
                catch (System.Exception exp)
                {
                    string error = exp.Message.ToString();
                 
                    MessageBox.Show("Failed to Update:" + error);
                }
                cn.Close();
        }

        private void ultraButton1_Click_1(object sender, EventArgs e)
        {
            txtPO.Text = "";
            txtLineNum.Text = "";
            clearALL();
        }

        private void DateTimePickerRelease_date_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                string v_site = this.txtLocation.Text.ToString().Replace("-BUILD", "");
                string v_part_num = this.txtPartNumber.Text.ToString();
                string vend_code = this.textBoxVendor.Text.ToString();
                string quarter = this.textBoxQuarter.Text.ToString();
                string ship_method_code = RepeaterPO.CurrentItem.Controls["comboBoxShipMethod"].Text;
                DateTime fact_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["DateTimePickerRelease_date"] as DateTimePicker).Value);
                string fob = this.textBoxFOB.Text;
                int days = getDueDate(v_site, vend_code, quarter, v_part_num, fob,ship_method_code);
                if (RadFactDate.Checked)
                {
                    //  (RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value = days;
                    (RepeaterPO.CurrentItem.Controls["dateTimePickerFacoryDate"] as DateTimePicker).Value = fact_date.AddDays(-1 * days);
                    //  (RepeaterPO.CurrentItem.Controls["dateTimePickerFacoryDate"] as DateTimePicker).Text = fact_date.AddDays(-1*days).ToString();
                }
            }
            catch (System.Exception exe)
            {
            }
        }

        string filePath = "\\\\Mvfile\\mv reports\\Master Data & Forms\\Application Forms\\";//atapps\\shopfloordocs\\CrystalReports\\";
        string strRptName = "adm_poform.rpt";
        private void ultraButton2_Click(object sender, EventArgs e)
        {
            int rev = 0;
            string sqlstr = " select m_revision from purchase where po_no='" + txtPO.Text.Trim() + "'";
            string constr = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;
            DataSet dsRev = getDataSet(sqlstr, constr);
            if (dsRev.Tables.Count > 0)
                if (dsRev.Tables[0].Rows.Count > 0)
                    rev = Convert.ToInt32(dsRev.Tables[0].Rows[0][0].ToString());

            ReportDocument report = new ReportDocument();
            String strReportName = filePath + strRptName;
            report.Load(strReportName);
            string PO_NUM = txtPO.Text.Trim();
            report.SetParameterValue("PO NUM", PO_NUM);
            //  CrystalReportViewer1.Refresh();
            //     CrystalReportViewer1.ReportSource = report;
            //    CrystalDecisions.Shared.PdfRtfWordFormatOptions opt=new CrystalDecisions.Shared.PdfRtfWordFormatOptions();

            // report.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.WordForWindows, "c:\\abc.pdf");
            //     CrystalReportViewer1.
            //  CrystalReportViewer1.ExportReport();
            Exportfilename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + PO_NUM + "_R" + rev + ".pdf";

            try
            {
                ExportOptions CrExportOptions = new ExportOptions();
                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();
                CrDiskFileDestinationOptions.DiskFileName = Exportfilename;// "E:\\crystalExportPO.pdf";
                CrExportOptions = report.ExportOptions;
                CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                CrExportOptions.FormatOptions = CrFormatTypeOptions;
                report.Export();
                
                sendEMail(rev);
            }
            catch (System.Exception ex)
            {
               // MessageBox.Show(ex.Message.ToString());

                //
            }


        }



        private void sendEMail(int rev)
        {
            Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.MailItem oMailItem = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            oMailItem.Attachments.Add(Exportfilename);
            // body, bcc etc...
            oMailItem.Subject = "PO NUM : " + txtPO.Text.Trim() + "  REV" + rev.ToString();
            oMailItem.Display(true);
        }

        private void ultraDateTimeFactoryDate_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                CheckBox chkConfirm = RepeaterPO.CurrentItem.Controls["checkBoxConfirmed"] as CheckBox;
                if (!chkConfirm.Checked)
                {
                    string v_site = this.txtLocation.Text.ToString().Replace("-BUILD", "");
                    string v_part_num = this.txtPartNumber.Text.ToString();
                    string vend_code = this.textBoxVendor.Text.ToString();
                    string quarter = this.textBoxQuarter.Text.ToString();
                    string ship_method_code = RepeaterPO.CurrentItem.Controls["comboBoxShipMethod"].Text;
                    DateTime fact_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeFactoryDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                    string fob = this.textBoxFOB.Text;
                    int days = getDueDate(v_site, vend_code, quarter, v_part_num,fob, ship_method_code);
                    if (RadMRPDate.Checked)
                    {
                        // (RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value = days;
                        (RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value = fact_date.AddDays(days);
                        //  (RepeaterPO.CurrentItem.Controls["DateTimePickerRelease_date"] as DateTimePicker).Text = fact_date.AddDays(days).ToString();
                    }

                }
            }
            catch (System.Exception ex)
            {
                string excepti = ex.Message.ToString();
            }

        }

        private void ultraDateTimeMRPDate_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                string v_site = this.txtLocation.Text.ToString().Replace("-BUILD", "");
                string v_part_num = this.txtPartNumber.Text.ToString();
                string vend_code = this.textBoxVendor.Text.ToString();
                string quarter = this.textBoxQuarter.Text.ToString();
                string ship_method_code = RepeaterPO.CurrentItem.Controls["comboBoxShipMethod"].Text;
                DateTime fact_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                string fob = this.textBoxFOB.Text;
                int days = getDueDate(v_site, vend_code, quarter, v_part_num,fob, ship_method_code);
                (RepeaterPO.CurrentItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value = days;
                if (RadFactDate!=null)
                if (RadFactDate.Checked)
                {
                    //   if (Convert.ToInt32((RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value) != Convert.ToInt32((RepeaterPO.CurrentItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value))
                    //       (RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Appearance.BackColor = System.Drawing.Color.Orange;
                    //    else
                    //        (RepeaterPO.CurrentItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Appearance.BackColor = System.Drawing.Color.White;

                    (RepeaterPO.CurrentItem.Controls["ultraDateTimeFactoryDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value = fact_date.AddDays(-1 * days);
                    //  (RepeaterPO.CurrentItem.Controls["dateTimePickerFacoryDate"] as DateTimePicker).Text = fact_date.AddDays(-1*days).ToString();
                }
            }
            catch (System.Exception ex)
            {
                string excepti = ex.Message.ToString();
            }

            try
            {
                DateTime mrp_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);

                //  string confirmquarter = 'Q' + this.GetQuarterName(mrp_date).ToString() + "-" + mrp_date.Year.ToString();
                CheckBox chkStatus = RepeaterPO.CurrentItem.Controls["checkBoxStatus"] as CheckBox;
                if (!chkStatus.Checked)
                {
                    if (mrp_date < System.DateTime.Now.Date)
                    {
                        RepeaterPO.CurrentItem.Controls["textBoxWarning"].Text = "*Past Due";
                    }

                    else
                    {
                        RepeaterPO.CurrentItem.Controls["textBoxWarning"].Text = "";
                    }
                }
            }
            catch (System.Exception ex)
            {
                string excepti = ex.Message.ToString();
            }

            try{

            DateTime mrp_date_value = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
              CheckBox chkStatus = RepeaterPO.CurrentItem.Controls["checkBoxStatus"] as CheckBox;
              DateTime request_date = System.DateTime.Now;
              if (chkStatus.Checked)
              {
                request_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeConfirmDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
              }
              else
              {
                  request_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeFactoryDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
              }
     
            TimeSpan ts = Convert.ToDateTime(mrp_date_value) - Convert.ToDateTime(request_date);
            if (ts.TotalDays == 0)
                RepeaterPO.CurrentItem.Controls["time_lead"].Text = "0";
            else
            {
                RepeaterPO.CurrentItem.Controls["time_lead"].Text = ts.TotalDays.ToString();
            }
            if (Convert.ToInt32(RepeaterPO.CurrentItem.Controls["time_lead"].Text) != Convert.ToInt32((RepeaterPO.CurrentItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value))
            {
                //  (e.DataRepeaterItem.Controls["time_lead"] as TextBox).Text = dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["diff"].ToString();
                RepeaterPO.CurrentItem.Controls["time_lead"].ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                //  (e.DataRepeaterItem.Controls["time_lead"] as TextBox).Text = dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["diff"].ToString();
                RepeaterPO.CurrentItem.Controls["time_lead"].ForeColor = System.Drawing.Color.Black;
            }
            }
            catch (System.Exception ex)
            {
                string excepti = ex.Message.ToString();
            }
        }

        private void checkBoxConfirmed_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                CheckBox chkConfirm = RepeaterPO.CurrentItem.Controls["checkBoxConfirmed"] as CheckBox;
                if (chkConfirm.Checked)
                {
                    RepeaterPO.CurrentItem.Controls["ultraDateTimeConfirmDate"].Visible = true;
                }
                else
                {
                    RepeaterPO.CurrentItem.Controls["ultraDateTimeConfirmDate"].Visible = false;
                }
            }
            catch (System.Exception ex)
            {
                string ee = ex.Message.ToString();
            }
            getTimeLead();
       
            
        }

        private void ultraDateTimeConfirmDate_ValueChanged(object sender, EventArgs e)
        {
            try
            {

                DateTime confirm_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeConfirmDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                string confirmquarter = 'Q' + this.GetQuarterName(confirm_date).ToString() + "-" + confirm_date.Year.ToString();
                //if (confirm_date < System.DateTime.Now.Date)
                //{
                //    RepeaterPO.CurrentItem.Controls["textBoxWarning"].Text = "*Past Due";
                //}
                //else 
                //if (confirm_date != null)
                //{
                //    CheckBox chkConfirm = RepeaterPO.CurrentItem.Controls["checkBoxConfirmed"] as CheckBox;
                //    chkConfirm.Checked = true;
                //}

                CheckBox chkStatus = RepeaterPO.CurrentItem.Controls["checkBoxStatus"] as CheckBox;
                if (!chkStatus.Checked)
                {
                    if (confirmquarter != textBoxQuarter.Text.ToString())
                    {
                        RepeaterPO.CurrentItem.Controls["textBoxWarning"].Text = "*Commit Qtr <> PO Qtr";
                    }
                    else
                    {
                        RepeaterPO.CurrentItem.Controls["textBoxWarning"].Text = "";
                    }
                }
            }
            catch(System.Exception ex)
            {
                string excepti = ex.Message.ToString();
            }
            try
            {
                CheckBox chkConfirm = RepeaterPO.CurrentItem.Controls["checkBoxConfirmed"] as CheckBox;
                if (chkConfirm.Checked)
                {
                    string v_site = this.txtLocation.Text.ToString().Replace("-BUILD", "");
                    string v_part_num = this.txtPartNumber.Text.ToString();
                    string vend_code = this.textBoxVendor.Text.ToString();
                    string quarter = this.textBoxQuarter.Text.ToString();
                    string ship_method_code = RepeaterPO.CurrentItem.Controls["comboBoxShipMethod"].Text;
                    DateTime fact_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeConfirmDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                    string fob = this.textBoxFOB.Text;
                    int days = getDueDate(v_site, vend_code, quarter, v_part_num, fob,ship_method_code);

                    (RepeaterPO.CurrentItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value = days;
                    if (RadMRPDate.Checked)
                    {
                        (RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value = fact_date.AddDays(days);
                        //  (RepeaterPO.CurrentItem.Controls["DateTimePickerRelease_date"] as DateTimePicker).Text = fact_date.AddDays(days).ToString();
                    }
                }
            }
            catch (System.Exception ex)
            {
                string excepti = ex.Message.ToString();
            }
            getTimeLead();
        }


        public void getTimeLead() 

        {

            try
            {
                DateTime request_date = System.DateTime.Now;

                DateTime mrp_date_value = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeMRPDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value);
                if ((RepeaterPO.CurrentItem.Controls["checkBoxConfirmed"] as CheckBox).Checked)

                    request_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeConfirmDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value).Date;
                else
                {
                    request_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["ultraDateTimeFactoryDate"] as Infragistics.Win.UltraWinEditors.UltraDateTimeEditor).Value).Date; 
                }


                TimeSpan ts = Convert.ToDateTime(mrp_date_value) - Convert.ToDateTime(request_date);
                if (ts.TotalDays == 0)
                    RepeaterPO.CurrentItem.Controls["time_lead"].Text = "0";
                else
                {
                    RepeaterPO.CurrentItem.Controls["time_lead"].Text = ts.TotalDays.ToString();
                }

                if (Convert.ToInt32(RepeaterPO.CurrentItem.Controls["time_lead"].Text) != Convert.ToInt32((RepeaterPO.CurrentItem.Controls["txtLeadTime"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value))
                {
                    //   (e.DataRepeaterItem.Controls["time_lead"] as TextBox).Text = dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["diff"].ToString();
                    RepeaterPO.CurrentItem.Controls["time_lead"].ForeColor = System.Drawing.Color.Red;
                }
                else
                {
                    //   (e.DataRepeaterItem.Controls["time_lead"] as TextBox).Text = dtRelease.Rows[e.DataRepeaterItem.ItemIndex]["diff"].ToString();
                    RepeaterPO.CurrentItem.Controls["time_lead"].ForeColor = System.Drawing.Color.Black;
                }
            }
            catch (System.Exception ex)
            { 
            
            }
    }

        private void txtQty_ValueChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    Infragistics.Win.UltraWinEditors.UltraNumericEditor txtqty = (RepeaterPO.CurrentItem.Controls["txtQty"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor);
            //    if ((Convert.ToInt32(txtqty.Text) != 0) && Convert.ToInt32(textBoxOrderMult.Text) != 0)
            //    {
            //        if (Convert.ToInt32(txtqty.Text) % Convert.ToInt32(textBoxOrderMult.Text) != 0)
            //        {
            //            RepeaterPO.CurrentItem.Controls["textBoxWarning"].Text = "*Order qty <> Order Multiple";
            //        }
            //        else
            //        {
            //            RepeaterPO.CurrentItem.Controls["textBoxWarning"].Text = "";
            //        }
            //    }
            //}
            //catch (System.Exception exp) 
            //{
            //}

        }

        private void txtQty_Leave(object sender, EventArgs e)
        {
            txtQty.Value = txtQty.Text;
        }
        DataTable dtTracker;
        private void RepeaterPO_CurrentItemIndexChanged(object sender, EventArgs e)
        {
            loadTracker();
        }

        private void loadTracker()
        {
            try
            {               
                int row_id = 0;
                this.txtTrackTotQty.Value = 0;
                if (RepeaterPO.CurrentItem.Controls["textBoxRowID"].Text.ToString() != "")
                {
                    row_id = Convert.ToInt32(RepeaterPO.CurrentItem.Controls["textBoxRowID"].Text.ToString());
                    lblRowID.Text = " Row ID :" + row_id.ToString();
                    string sqlTracker = " SELECT [id]     ,[Qty]      ,[Carrier]      ,[Tracking], delivered,notes  FROM [MIMDIST].[dbo].[releases_tracking] where row_id=" + row_id;
                    sqlTracker = sqlTracker + "   SELECT isnull(sum([Qty]),0)  FROM [MIMDIST].[dbo].[releases_tracking] where row_id=" + row_id;
                    //if  ( this.txtVendor.Text.Trim() != "")
                    //    sql = sql + " and vendor_no='" + this.txtVendor.Text.Trim() + "'";

                    this.repeaterTracker.DataSource = null;
                    string constr = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;
                    DataSet dsTracker = getDataSet(sqlTracker, constr);
                    dtTracker = dsTracker.Tables[0];//.Select("po_line=1").CopyToDataTable();
                    bindingSourceTracker.DataSource = dtTracker;
                    this.repeaterTracker.DataSource = bindingSourceTracker;
                    if ((RepeaterPO.CurrentItem.Controls["checkBoxStatus"] as CheckBox).Checked)
                        repeaterTracker.Enabled = false;
                    else
                        repeaterTracker.Enabled = true;
                    this.txtTrackTotQty.Value = 0;
                    if (dsTracker.Tables[1] != null)
                        if (dsTracker.Tables[1].Rows.Count > 0)
                            this.txtTrackTotQty.Value = Convert.ToInt32(dsTracker.Tables[1].Rows[0][0].ToString());
                    //releasesBindingSource.DataSource = dtTracker;
                }
                else
                {

                    lblRowID.Text = "";
                   // pnlDetail.Visible = false;
                }
            }
            catch(System.Exception e) {

               // MessageBox.Show(e.Message);
            }
        }
          
        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                repeaterTracker.AddNew();
               // bindingSourceTracker.AddNew();
            }
            catch
            {
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dtTracker.Rows.Count > 0)
            {
                DeleteTrackerLine();
            }
            else
            {
                MessageBox.Show("Please select the row");
            }
            loadTracker();
            bindgrid();
        }

        public void DeleteTrackerLine()
        {
          int  row_id = Convert.ToInt32(lblRowID.Text.Replace("Row ID :", "").Trim());
            if (repeaterTracker.CurrentItem.Controls["txtID"].Text.Trim() != "")
            {
                string sql = "delete from releases_tracking where id=" + repeaterTracker.CurrentItem.Controls["txtID"].Text + " select @@error    ";
                string constr = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;
                DataSet ds = getDataSet(sql, constr);
                if (ds.Tables[0].Rows[0][0].ToString() == "0")
                {
                    MessageBox.Show("Item Deleted");
                    string qry = "declare @cnt int select @cnt=count(*) from dbo.releases_tracking where row_id= " + row_id.ToString() + "   if (@cnt>0)   update releases set m_ship_confirm='Y' where row_id= " + row_id.ToString() + "   else  update releases set m_ship_confirm='N' where row_id= " + row_id.ToString();
                    DataSet ds1 = getDataSet(qry, constr);
                }
                else
                {
                    MessageBox.Show("Error to delete Item");
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            int row_id = 0;
            string v_carrier="";
            string v_m_tracking="";
            int v_Tracking_qty=0;
            string v_tracking_num="";
            string insertqry="";
            string updateqry="";
            int v_id=0;
            int total_qty = 0;
            bool v_delivered;
             string notes="";
             bool NoSave = false;
              row_id = Convert.ToInt32(lblRowID.Text.Replace("Row ID :","").Trim());
              if (row_id != 0)
              {
                  for (int i = 0; i < repeaterTracker.ItemCount; i++)
                  {
                      repeaterTracker.CurrentItemIndex = i;
                      Microsoft.VisualBasic.PowerPacks.DataRepeaterItem item = repeaterTracker.CurrentItem;
                      if (repeaterTracker.CurrentItem.Controls["txtID"].Text.ToString() != "")
                          v_id = Convert.ToInt32(repeaterTracker.CurrentItem.Controls["txtID"].Text.ToString());
                      else
                          v_id = 0;

                      /*   if (item.Controls["comboBoxCarrier"].Text.ToString() != "")
                         {
                             v_carrier = dtCarrier.Select("ship_via_name='" + item.Controls["comboBoxCarrier"].Text.ToString() + "'")[0][0].ToString();
                         }*/
                      v_m_tracking = repeaterTracker.CurrentItem.Controls["textBoxTracking"].Text.Trim();

                      v_Tracking_qty = Convert.ToInt32((repeaterTracker.CurrentItem.Controls["textBoxQTY"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value);

                      CheckBox chkDeliverred =  repeaterTracker.CurrentItem.Controls["chkDelelivered"] as CheckBox; //RepeaterPO.CurrentItem.Controls["checkBoxStatus"] as CheckBox;
                        notes= (this.repeaterTracker.CurrentItem.Controls["Note2"] as Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor).Text ;

                      // v_due_date = getDueDate(v_site, vend_code, quarter, v_part_num, v_m_factory_date, v_m_ship_method);
                      total_qty = total_qty + v_Tracking_qty;
                      bool regexmatch=true;
                      if (v_m_tracking !="")
                      {
                          string str = dtCarrier.Select("ship_via_name='" + RepeaterPO.CurrentItem.Controls["comboBoxCarrier"].Text + "'")[0][2].ToString();
                          if (str != "0")
                          {
                              System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(str);
                              regexmatch = regex.Match(v_m_tracking).Success;
                          }

                      }
                     

                           if (regexmatch)
                           {
                               if (v_Tracking_qty > 0)
                               {
                                   if (v_id == 0)
                                   {
                                       //insert into releases (po_no,part_no,location,part_type,release_date,quantity,received,status,confirm_date,confirmed,lb_tracking,conv_factor,prev_qty,po_key,due_date,ord_line,po_line,m_ship_method,m_factory_date,m_ship_confirm,m_tracking,carrier) ";
                                       insertqry = insertqry + "insert into   dbo.releases_tracking ( row_id, Qty, Tracking,delivered,notes) values(" + row_id + "," + v_Tracking_qty + ",'" + v_m_tracking + "','" + chkDeliverred.Checked.ToString() + "','" + notes + "')";
                                   }
                                   else
                                   {
                                       updateqry = updateqry + " update dbo.releases_tracking set  Qty=" + v_Tracking_qty + ", Tracking='" + v_m_tracking + "', delivered='" + chkDeliverred.Checked.ToString() + "',notes='" + notes + "',[modified_date]=getdate() where id=" + v_id.ToString();
                                   }
                                   repeaterTracker.CurrentItem.Controls["textBoxTracking"].BackColor = System.Drawing.Color.White;
                               }
                               else
                               {
                                   MessageBox.Show("Quantity must be greater than zero.");
                               }
                           }
                           else
                           {
                               repeaterTracker.CurrentItem.Controls["textBoxTracking"].BackColor = System.Drawing.Color.Red;
                               NoSave = true;
                           }
                                           
                  }

                  int Release_qty = 0;
                  if (RepeaterPO.CurrentItem.Controls["txtQty"].Text.Trim() != "")
                      Release_qty = Convert.ToInt32((RepeaterPO.CurrentItem.Controls["txtQty"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value);

                  if (NoSave)
                  {
                      MessageBox.Show("Regex for tracking number doesn't match.");
                  }
                  else if (v_Tracking_qty > 0)
                  {
                      if ((total_qty == Release_qty) )
                      {

                          string qry = insertqry + updateqry + "declare @cnt int select @cnt=count(*) from dbo.releases_tracking if (@cnt>0)   update releases set m_ship_confirm='Y' where row_id= " + row_id.ToString() + "   else  update releases set m_ship_confirm='N' where row_id= " + row_id.ToString()  ;
                          string constr = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;
                          SqlConnection cn = new SqlConnection(constr);
                          cn.Open();
                          SqlTransaction tran = cn.BeginTransaction("Transaction");
                          SqlCommand cmd = new SqlCommand(qry, cn);
                          cmd.Transaction = tran;
                          try
                          {
                              SqlDataAdapter da = new SqlDataAdapter(cmd);
                              DataSet ds = new DataSet();
                              da.Fill(ds);
                              tran.Commit();
                              MessageBox.Show("Updated successfully");
                            

                          }
                          catch (System.Exception exp)
                          {
                              string error = exp.Message.ToString();
                              tran.Rollback();
                              MessageBox.Show("Failed to Update:" + error);
                          }
                          cn.Close();
                          loadTracker();
                          bindgrid();
                      }
                      else
                      {
                          MessageBox.Show("Tracker can not be saved! Release quantity doesn't match with shippping quantity");
                      }
                  }
              }
        }
            
   
        private void repeaterTracker_DrawItem(object sender, Microsoft.VisualBasic.PowerPacks.DataRepeaterItemEventArgs e)
        {
          /*  try
            {
                ComboBox CmbCarrier = e.DataRepeaterItem.Controls["comboBoxCarrier"] as ComboBox;
                try
                {
                    if (dtTracker != null)
                        for (int i = 0; i < this.dtCarrier.Rows.Count; i++)
                        {
                            // if (dtCarrier.Rows.Count > e.DataRepeaterItem.ItemIndex)
                            // {
                            string tracker_carrier = dtCarrier.Rows[i]["ship_via_code"].ToString();
                            string carrier_string = dtTracker.Rows[e.DataRepeaterItem.ItemIndex]["carrier"].ToString();
                            if (tracker_carrier == carrier_string)
                                CmbCarrier.SelectedIndex = i;
                            // }
                        }
                }
                catch
                {
                    CmbCarrier.SelectedIndex = 0;
                }
            }
            catch {
                
            }*/
        

         


        }

        private void repeaterTracker_ItemCloned(object sender, Microsoft.VisualBasic.PowerPacks.DataRepeaterItemEventArgs e)
        {
          /*  try
            {
                string constr = ConfigurationManager.ConnectionStrings["epicorConnectionStringData"].ConnectionString;
                ComboBox CmbCarrier = e.DataRepeaterItem.Controls["comboBoxCarrier"] as ComboBox;

                dtCarrier = getDataSet("SELECT ship_via_code, ship_via_name FROM arshipv", constr).Tables[0];
                if (CmbCarrier.DataSource == null)
                {
                    CmbCarrier.DisplayMember = "ship_via_name";
                    CmbCarrier.ValueMember = "ship_via_code";
                    CmbCarrier.DataSource = dtCarrier;
                }
                CmbCarrier.SelectedIndexChanged += new EventHandler(comboBox2_SelectedIndexChanged);
            }
            catch { }*/
        }

        private void textBoxQTY_Leave(object sender, EventArgs e)
        {
           textBoxQTY.Value = textBoxQTY.Text;
        }
        public string save_int_notes = "";
        public string save_ext_notes = "";
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            NotePad note = new NotePad(internal_notes,txtPO.Value.ToString());
            note.Text = "INTERNAL NOTES";
            note.ShowDialog();          
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            NotePad note = new NotePad(external_notes, txtPO.Value.ToString());
            note.Text = "EXTERNAL NOTES";
            note.ShowDialog();
        }

        private void txtTrackTotQty_ValueChanged(object sender, EventArgs e)
        {
           
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string note = (this.repeaterTracker.CurrentItem.Controls["Note2"] as Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor).Text;
            NotePad note1 = new NotePad(note, "");
            note1.Text = "TRACKING NOTE";
           
            note1.ShowDialog();
            (this.repeaterTracker.CurrentItem.Controls["Note2"] as Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor).Text = note1.notes;
        }

        private void RepeaterPO_Click(object sender, EventArgs e)
        {
           // MessageBox.Show("here");
            if (RepeaterPO.CurrentItemIndex >= 0)
                loadTracker();
        }

        private void picReleaseNote_Click(object sender, EventArgs e)
        {
            string note = (this.RepeaterPO.CurrentItem.Controls["txtReleaseNote"] as Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor).Text;
            NotePad note1 = new NotePad(note, "");
            note1.Text = "RELEASE NOTE";
            note1.ShowDialog();
            (this.RepeaterPO.CurrentItem.Controls["txtReleaseNote"] as Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor).Text = note1.notes;
        }

        private void time_lead_ValueChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    string v_site = this.txtLocation.Text.ToString().Replace("-BUILD", "");
            //    string v_part_num = this.txtPartNumber.Text.ToString();
            //    string vend_code = this.textBoxVendor.Text.ToString();
            //    string quarter = this.textBoxQuarter.Text.ToString();
            //    var combo = (Infragistics.Win.UltraWinEditors.UltraNumericEditor) sender;
            //    var DataRepeaterItem = (Microsoft.VisualBasic.PowerPacks.DataRepeaterItem)combo.Parent;
            //    var comboship = (DataRepeaterItem.Controls["comboBoxShipMethod"] as ComboBox);
            //    string ship_method_code = comboship.Text;
            //    // DateTime fact_date = Convert.ToDateTime((RepeaterPO.CurrentItem.Controls["dateTimePickerFacoryDate"] as DateTimePicker).Text);
            //    int days = getDueDate(v_site, vend_code, quarter, v_part_num, ship_method_code);
            //    if (Convert.ToInt32((DataRepeaterItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Value) != days)
            //    {
            //        (DataRepeaterItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Appearance.BackColor = System.Drawing.Color.Orange;
            //    }
            //    else
            //    {
            //        (DataRepeaterItem.Controls["time_lead"] as Infragistics.Win.UltraWinEditors.UltraNumericEditor).Appearance.BackColor = System.Drawing.Color.White;
            //    }
            //}

            //catch { }


        }

        private void ultraNumericEditor1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void comboBoxShipMethod_SelectedIndexChanged_2(object sender, EventArgs e)
        {

        }

        private void time_lead_TextChanged(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

      
      
    }

    
}

   
