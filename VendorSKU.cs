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

namespace Version3
{
    public partial class VendorSKU : Form
    {
        string constrMVErp = ConfigurationManager.ConnectionStrings["mverpConnectionString"].ConnectionString;
        public VendorSKU()
        {
            InitializeComponent();           
        }

        public DataSet getDataSet(string sqlquery,string constr)
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
        
        public void bindGrid()
        {
            string constr = ConfigurationManager.ConnectionStrings["epConnectionString"].ConnectionString;
            string strqry = "SELECT repl_id,vend.vend_code as Vendor ,vend.mim_pn as [Part Num],mfg_pn as [Manf Part Num],last_price as [Price],curr_key as Currency,mfg_lead_time as [Manf Lead Time],quote_date as [Quote Date] ,fob_at as [FOB AT],fob_nl as [FOB NL],ship_meth_at  as [Ship Meth AT],ship_meth_nl  as [Ship Meth NL],transit_lt_at as [AT Transit LT],transit_lt_nl as [NL Transit LT],default_carrier as [Default Carr],[qty_per_box] as [Box QTY],[boxes_per_pallet] as [Pallet QTY],ord_mult as [Order Multiple],quote_qtr FROM dbo.m_app_vendor_sku_view vend where vend_code<>''";

            //  string strqry = "SELECT vendor_sku.m_quote_date, vendor_sku.vendor_no, vendor_sku.sku_no, inv_master.description, vendor_sku.curr_key, vendor_sku.vend_sku, vendor_sku.last_price, vendor_sku.qty, vendor_sku.m_manuf_lead_time, vendor_sku.fob_at, vendor_sku.ship_meth_at, vendor_sku.fob_nl,vendor_sku.ship_meth_nl, vendor_sku.fob_sg, vendor_sku.ship_meth_sg, apmaster.address_name FROM inv_master INNER JOIN vendor_sku ON inv_master.part_no = vendor_sku.sku_no INNER JOIN apmaster ON vendor_sku.vendor_no = apmaster.vendor_code WHERE (apmaster.status_type = 5) AND (apmaster.vend_class_code = 'VENDOR')";
            string vendor_no = vendor.Text.Trim();
            string qtr = this.cmbQtrFilter.Text.Trim();
            string partNum = this.txtPartNumber.Text.Trim();
            if (vendor_no != "")
            {
                strqry = strqry + "  and    vend_code='" + vendor_no + "'";
            }
            if (qtr != "")
                    strqry = strqry + " and quote_qtr='" + qtr + "' ";

                if (partNum != "")
                    strqry = strqry + " and vend.mim_pn like '" + partNum + "%' ";

                strqry = strqry + "  order by vend.mim_pn ";
                SqlConnection cn = new SqlConnection(constr);
                cn.Open();
                SqlCommand cmd = new SqlCommand(strqry, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                cn.Close();
                vendorGrid.DataSource = ds.Tables[0];
                bindingSource1.DataSource = ds.Tables[0];

            //}
        }

        private void ultraButton1_Click(object sender, EventArgs e)
        {

           
        }

        private void ultraButton2_Click(object sender, EventArgs e)
        {
            this.vendorGrid.PerformAction(UltraGridAction.Copy);
        }

        private void ultraButton3_Click(object sender, EventArgs e)
        {      
                       
        }
        

        public void bindVendor()
        {
            string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
            string strqry = " SELECT distinct vendor_code FROM [MIMDIST].[dbo].[apmaster] where status_type=5 and [vend_class_code]='VENDOR'  SELECT distinct part_no FROM [MIMDIST].[dbo].[inv_master] WHERE part_no=sku_no and    (void = 'N')  AND (obsolete = 0)";
            DataSet ds = getDataSet(strqry,constr);
            DataTable dtVendor= ds.Tables[0] ;

            DataRow dr = dtVendor.NewRow();
            dr[0] = "";
            dtVendor.Rows.InsertAt(dr, 0);

            this.vendor.DataSource = dtVendor;
            vendor.ValueMember = "vendor_code";
            vendor.DisplayMember = "vendor_code";            
            this.comboPart.DataSource = ds.Tables[1];
            comboPart.ValueMember = "part_no";
            comboPart.DisplayMember = "part_no";            
            comboVendor.DataSource = ds.Tables[0];
            comboVendor.ValueMember = "vendor_code";
            comboVendor.DisplayMember = "vendor_code";
        }

        public void bindCurrency()
        {
            DataTable dtCurrency = new DataTable();
            dtCurrency.Columns.Add(new DataColumn("Currency"));
            DataRow dr = dtCurrency.NewRow();
            dr[0] = "EUR";
            dtCurrency.Rows.Add(dr);
            dr = dtCurrency.NewRow();
            dr[0] = "USD";
            dtCurrency.Rows.Add(dr);
            comboCurrency.DataSource = dtCurrency;
            comboCurrency.ValueMember = "Currency";
            comboCurrency.DisplayMember = "Currency";
        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            this.vendorGrid.DisplayLayout.Bands[0].AddNew();
            this.vendorGrid.Rows[vendorGrid.Rows.Count - 1].Selected = true;
        }


        string rep_id_string = "";

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {     
             DialogResult    myResult = MessageBox.Show("Are you really delete the item?", "Delete Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
             if (myResult == DialogResult.OK)
             {

                 DataTable dt = (this.vendorGrid.DataSource as DataTable).Copy();
                 //  RowsCollection rc ;
                 foreach (UltraGridRow row in vendorGrid.Rows)
                 {
                     // MessageBox.Show(row.Cells[row.Cells.Count - 1].Text.ToString());
                     if (row.Cells[row.Cells.Count - 1].Text.ToString() == "True")
                     {
                         //  row.Delete();
                         DataRow[] drRep = dt.Select("repl_id=" + row.Cells["repl_id"].Value.ToString() + " and [part num]='" + row.Cells["part num"].Value.ToString() + "'");
                         rep_id_string = rep_id_string + "," + row.Cells["repl_id"].Value.ToString();
                         if (drRep.Length > 0)
                             dt.Rows.Remove(drRep[0]);
                     }
                 }
                 vendorGrid.DataSource = dt;
             }
             string qryDelete = "";
             if (rep_id_string != "")
             {
                 qryDelete = "  Delete from dbo.vendor_sku where repl_id in( " + rep_id_string.Substring(1, rep_id_string.Length - 1) + ")";
             }
             string constr = ConfigurationManager.ConnectionStrings["epConnectionString"].ConnectionString;
     
             SqlConnection cn = new SqlConnection(constr);
             cn.Open();
             SqlTransaction tran = cn.BeginTransaction("Transaction");
             SqlCommand cmd = new SqlCommand(qryDelete, cn);
             cmd.Transaction = tran;
             try
             {
                 SqlDataAdapter da = new SqlDataAdapter(cmd);
                 DataSet ds = new DataSet();
                 da.Fill(ds);
                 tran.Commit();
                 MessageBox.Show("Deleted successfully");

             }
             catch (Exception exp)
             {
                 string error = exp.Message.ToString();
                 tran.Rollback();
                 MessageBox.Show("Failed to Update. Please check the data");
             }
             cn.Close();
             bindGrid(); 
             
        }

        private void ugData_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e)
        {
            e.Cancel = true;
        }

        private void copyToolStripButton_Click(object sender, EventArgs e)
        {
            DataTable dt = (this.vendorGrid.DataSource as DataTable).Copy();
            foreach (UltraGridRow row in this.vendorGrid.Rows)
            {
                if (row.Cells[row.Cells.Count - 1].Text.ToString() == "True")
                {                 
                    DataRow dr=dt.NewRow();
                    dr[0] = 0;
                    for (int i = 1; i < row.Cells.Count-1; i++)
                        {
                            if (dt.Columns[i].ColumnName == "quote_qtr")
                                dr[i] = "";
                            else if (dt.Columns[i].ColumnName == "Price")
                                dr[i] = "0";
                            else
                                dr[i] = row.Cells[i].Value;
                           
                        }
                    dt.Rows.Add(dr);
                }
            }               
            vendorGrid.DataSource = dt;
       //     this.vendorGrid.PerformAction(UltraGridAction.Copy);
       //     this.vendorGrid.DisplayLayout.Bands[0].AddNew();
       //     this.vendorGrid.Rows[vendorGrid.Rows.Count - 1].Selected = true;
       //     this.vendorGrid.PerformAction(UltraGridAction.Paste);
        }


        private void pasteToolStripButton_Click(object sender, EventArgs e)
        {
            this.vendorGrid.PerformAction(UltraGridAction.Paste);
        }

        private void ultraButton1_Click_1(object sender, EventArgs e)
        {
            bindGrid();
            loadComboData();
        }
              

        private void VendorSKU_Load(object sender, EventArgs e)
        {
            BindUltraDropDown();
            bindVendor();
            bindQrter();
            bindQrterFilter();
            bindCurrency();
            loadCarrier();
        }
        
        private void BindUltraDropDown()
        {
            DataTable dt = new DataTable();
            //  dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("fob", typeof(string));
            dt.Rows.Add(new object[] { "US" });
            dt.Rows.Add(new object[] { "EUROPE" });
            dt.Rows.Add(new object[] { "ASIA" });
            dt.AcceptChanges();

            this.comboFOB.SetDataBinding(dt, null);
            this.comboFOB.ValueMember = "fob";
            this.comboFOB.DisplayMember = "fob";

            DataTable dt1 = new DataTable();
            //  dt.Columns.Add("ID", typeof(int));
            dt1.Columns.Add("ship meth", typeof(string));

            dt1.Rows.Add(new object[] { "GROUND" });
            dt1.Rows.Add(new object[] { "STANDARD AIR" });
            dt1.Rows.Add(new object[] { "OCEAN" });
            dt1.AcceptChanges();

            this.comboShipMethod.SetDataBinding(dt1, null);
            this.comboShipMethod.ValueMember = "ship meth";
            this.comboShipMethod.DisplayMember = "ship meth";
        }


        public void bindQrter()
        {
          //  string sql = " SELECT 'Q'+DATENAME(Quarter, GETDATE())+'-'+cast(YEAR(getdate()) as varchar) as Quarter,1 as ord union SELECT 'Q'+DATENAME(Quarter,DATEADD(QQ,1, GETDATE()))+'-'+ cast(YEAR(DATEADD(QQ,1, GETDATE())) as varchar) as Quarter ,2 as ord  union SELECT 'Q'+DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))+'-'+ cast(YEAR(DATEADD(QQ,2, GETDATE())) as varchar) as Quarter,3 as ord  order by  ord ";
          //  string sql="  SELECT cast(YEAR(getdate()) as varchar)+'Q'+replicate('0', (2-len(DATENAME(Quarter, GETDATE()))))+ DATENAME(Quarter, GETDATE()) as Quarter, 1 as ord   union  SELECT cast(YEAR(DATEADD(QQ,1, GETDATE())) as varchar) +'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,1, GETDATE())))))+DATENAME(Quarter,DATEADD(QQ,1, GETDATE())) as Quarter ,2 as ord   union   SELECT cast(YEAR(DATEADD(QQ,2, GETDATE())) as varchar)+'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))))) +DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))  as Quarter,3 as ord  order by  ord  ";
           // string sql = "  SELECT cast(YEAR(getdate()) as varchar)+'Q'+replicate('0', (2-len(DATENAME(Quarter, GETDATE()))))+ DATENAME(Quarter, GETDATE()) as Quarter   union  SELECT cast(YEAR(DATEADD(QQ,1, GETDATE())) as varchar) +'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,1, GETDATE())))))+DATENAME(Quarter,DATEADD(QQ,1, GETDATE())) as Quarter   union   SELECT cast(YEAR(DATEADD(QQ,2, GETDATE())) as varchar)+'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))))) +DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))  as Quarter  ";

            string sql = "SELECT TOP 1000 [Quarter]   FROM [MIMDIST].[dbo].[v_qtr] ";
            DataTable dtQuarter = getDataSet(sql).Tables[0];

            /*  
             * 
             *    dtQuarter.Columns.Add(new DataColumn("Quarter"));
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
                dtQuarter.Columns.Add(new DataColumn("Quarter"));
                DataRow dr = dtQuarter.NewRow();
                dr[0] = "Q1-2014";
                dtQuarter.Rows.Add(dr);
                dr = dtQuarter.NewRow();
                dr[0] = "Q2-2014";
                dtQuarter.Rows.Add(dr);
                dr = dtQuarter.NewRow();
                dr[0] = "Q3-2014";
                dtQuarter.Rows.Add(dr);
                dr = dtQuarter.NewRow();
                dr[0] = "Q4-2014";
                dtQuarter.Rows.Add(dr);
         
            */
          //  ComboQtr.DataSource = dtQuarter;
          //  ComboQtr.ValueMember = "Quarter";
          //  ComboQtr.DisplayMember = "Quarter";
            this.combogridQTR.DataSource = dtQuarter;
            combogridQTR.ValueMember = "Quarter";
            combogridQTR.DisplayMember = "Quarter";

            foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in combogridQTR.Rows)
            {
                if (row.Cells[0].Value.ToString() == currentqtr())
                {
                    combogridQTR.SelectedRow = row;
                }
            }


        }

        public void bindQrterFilter()
        {
            //  string sql = " SELECT 'Q'+DATENAME(Quarter, GETDATE())+'-'+cast(YEAR(getdate()) as varchar) as Quarter,1 as ord union SELECT 'Q'+DATENAME(Quarter,DATEADD(QQ,1, GETDATE()))+'-'+ cast(YEAR(DATEADD(QQ,1, GETDATE())) as varchar) as Quarter ,2 as ord  union SELECT 'Q'+DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))+'-'+ cast(YEAR(DATEADD(QQ,2, GETDATE())) as varchar) as Quarter,3 as ord  order by  ord ";
           // string sql = "  SELECT cast(YEAR(getdate()) as varchar)+'Q'+replicate('0', (2-len(DATENAME(Quarter, GETDATE()))))+ DATENAME(Quarter, GETDATE()) as Quarter   union  SELECT cast(YEAR(DATEADD(QQ,1, GETDATE())) as varchar) +'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,1, GETDATE())))))+DATENAME(Quarter,DATEADD(QQ,1, GETDATE())) as Quarter    union   SELECT cast(YEAR(DATEADD(QQ,2, GETDATE())) as varchar)+'Q'+replicate('0', (2-len(DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))))) +DATENAME(Quarter,DATEADD(QQ,2, GETDATE()))  as Quarter    ";
            string sql = "SELECT TOP 1000 [Quarter]   FROM [MIMDIST].[dbo].[v_qtr] ";
            DataTable dtQuarter = getDataSet(sql).Tables[0];
            DataRow drQtr = dtQuarter.NewRow();
            drQtr[0] = "";
            dtQuarter.Rows.InsertAt(drQtr,0);
            cmbQtrFilter.DataSource = dtQuarter;
            cmbQtrFilter.DisplayMember="Quarter";
            cmbQtrFilter.ValueMember = "Quarter";
            cmbQtrFilter.Value = currentqtr();
        }


        private string currentqtr()
        {
            string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
            string sql = " select dbo.udf_yearQTR(GETDATE())";
            return getDataSet(sql,constr).Tables[0].Rows[0][0].ToString();
        }

        static string ExtractNumbers(string expr)
        {
            return string.Join(null, System.Text.RegularExpressions.Regex.Split(expr, "[^\\d]"));
        }
        
        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;           
          
           // string m_qrt = ComboQtr.Text.Trim().ToString();
            string vendor_no = "";

            string qryDelete ="";
       //     if (rep_id_string!="")
        //    {
         //       qryDelete = "  Delete from dbo.vendor_sku where repl_id in( " + rep_id_string.Substring(1, rep_id_string.Length - 1) + ")";
          //  }
            string qry = "";
            string sku_list = "";
            bool duplicatePart = false;
            bool Qblank = false;
            sku_list = "";
            //MessageBox.Show(vendorGrid.Rows.Count.ToString());
            for (int i = 0; i < vendorGrid.Rows.Count; i++)
            {
                int rep_id = 0;
                string sku_no = vendorGrid.Rows[i].Cells["Part Num"].Text.ToString();
                string vendor = vendorGrid.Rows[i].Cells["Vendor"].Text.ToString();
                string sku_no_qtr = vendorGrid.Rows[i].Cells["Part Num"].Text.ToString() + '-' + vendorGrid.Rows[i].Cells["quote_qtr"].Text.ToString() + '-' + vendorGrid.Rows[i].Cells["Vendor"].Text.ToString();
                if (sku_list.Contains(sku_no_qtr))
                {
                    duplicatePart = true;
                   
                }
                sku_list = sku_list +"," + sku_no_qtr;
               //vendorGrid.Rows[i].Cells["vendor_no"].Value.ToString();
                string vend_sku = vendorGrid.Rows[i].Cells["Manf Part Num"].Text.ToString();
                string last_price = vendorGrid.Rows[i].Cells["Price"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString();
                int Gross_weight = 0;
                /*  if (vendorGrid.Rows[i].Cells["Gross_weight"].Text.ToString().Trim() != "")
                {
                    Gross_weight = Convert.ToInt32(vendorGrid.Rows[i].Cells["Gross_weight"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                }*/
                string curr_key = vendorGrid.Rows[i].Cells["Currency"].Text.ToString();
                int m_qty_per_box = 0;
                if (vendorGrid.Rows[i].Cells["Box QTY"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString().Trim() != "")
                {
                    m_qty_per_box = Convert.ToInt32(vendorGrid.Rows[i].Cells["Box QTY"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                }
                int m_boxes_per_pallet = 0;
                if (vendorGrid.Rows[i].Cells["Pallet QTY"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString().Trim() != "")
                {
                    m_boxes_per_pallet = Convert.ToInt32(vendorGrid.Rows[i].Cells["Pallet QTY"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                }

                vendor_no = vendorGrid.Rows[i].Cells["Vendor"].Text.ToString();
                string m_manuf_lead_time = vendorGrid.Rows[i].Cells["Manf Lead Time"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString();
              //  string m_trasit_lead_time = vendorGrid.Rows[i].Cells["m_trasit_lead_time"].Value.ToString();
                string m_quote_date = vendorGrid.Rows[i].Cells["Quote Date"].Text.ToString();

                string fob_at = vendorGrid.Rows[i].Cells["FOB AT"].Text.ToString();
                string fob_nl = vendorGrid.Rows[i].Cells["FOB NL"].Text.ToString();
              //  string fob_sg = vendorGrid.Rows[i].Cells["FOB SG"].Text.ToString();
                string ship_meth_at = vendorGrid.Rows[i].Cells["Ship Meth AT"].Text.ToString();
                string ship_meth_nl = vendorGrid.Rows[i].Cells["Ship Meth NL"].Text.ToString();
            //    string ship_meth_sg = vendorGrid.Rows[i].Cells["Ship Meth SG"].Text.ToString();

                string leadTime_at = vendorGrid.Rows[i].Cells["AT Transit LT"].Text.ToString();
                string leadTime_nl = vendorGrid.Rows[i].Cells["NL Transit LT"].Text.ToString();
                string m_qrt = vendorGrid.Rows[i].Cells["quote_qtr"].Text.ToString();
             //   string leadTime_sg = vendorGrid.Rows[i].Cells["SG Transit LT"].Text.ToString();
              int ord_mult = 0;
              if (vendorGrid.Rows[i].Cells["Order Multiple"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString().Trim() != "")
                {
                    ord_mult = Convert.ToInt32(vendorGrid.Rows[i].Cells["Order Multiple"].GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToString());
                }
               // string ord_mult = vendorGrid.Rows[i].Cells["Order Multiple"].Text.ToString();
                string default_carr = "";// dsComboCarrier.Tables[0].Select("ship_via_name='" + vendorGrid.Rows[i].Cells["Default Carr"].Text.ToString() + "'")[0]["ship_via_code"].ToString();
                DataRow[] drCarr =  dsComboCarrier.Tables[0].Select("ship_via_name='" + vendorGrid.Rows[i].Cells["Default Carr"].Text.ToString() + "'") as DataRow[];
                if (drCarr.Length>0)
                    default_carr =  dsComboCarrier.Tables[0].Select("ship_via_name='" + vendorGrid.Rows[i].Cells["Default Carr"].Text.ToString() + "'")[0]["ship_via_code"].ToString();
                if ((vendorGrid.Rows[i].Cells["repl_id"].Value != null) && (vendorGrid.Rows[i].Cells["repl_id"].Value.ToString() != ""))
                    rep_id = Convert.ToInt32(vendorGrid.Rows[i].Cells["repl_id"].Value.ToString());

                if (m_qrt == "")
                {
                    Qblank = true;
                }
                if (sku_no != "")
                {
                    //qry = qry + "    INSERT INTO dbo.vendor_sku(sku_no, vendor_no, vend_sku, last_price, m_gross_weight, curr_key, m_manuf_lead_time, m_quote_date, m_qrt, fob_at, fob_nl, fob_sg, ship_meth_at, ship_meth_nl, ship_meth_sg, m_qty_per_box,m_boxes_per_pallet,ord_mult) VALUES('" + sku_no + "','" + vendor_no + "','" + vend_sku + "','" + last_price + "','" + Gross_weight + "','" + curr_key + "','" + m_manuf_lead_time + "','" + m_quote_date + "','" + m_qrt + "','" + fob_at + "','" + fob_nl + "','" + fob_sg + "','" + ship_meth_at + "','" + ship_meth_nl + "','" + ship_meth_sg + "','" + m_qty_per_box + "','" + m_boxes_per_pallet + "','" + ord_mult + "')  ";
                    qry = qry + "    exec  sp_insert_update_vendor_sku    '" + sku_no + "','" + vendor_no + "','" + vend_sku + "','" + last_price  + "','" + curr_key + "','" + m_manuf_lead_time + "','" + m_quote_date + "','" + m_qrt + "','" + fob_at + "','" + fob_nl  + "','" + ship_meth_at + "','" + ship_meth_nl + "','" + m_qty_per_box + "','" + m_boxes_per_pallet + "','" + ord_mult + "','" + leadTime_at + "','" + leadTime_nl + "','"  + default_carr + "'," + rep_id;

                }

                   // if (rep_id == 0)
                   // {
                    //    qry = qry + "Delete from dbo.vendor_sku where vendor_no='" + vendor_no + "' and m_qrt='" + m_qrt + "'    INSERT INTO dbo.vendor_sku(sku_no, vendor_no, vend_sku, last_price, qty, curr_key, m_manuf_lead_time, m_trasit_lead_time, m_quote_date, m_qrt, fob_at, fob_nl, fob_sg, ship_meth_at, ship_meth_nl, ship_meth_sg, m_qty_per_box) VALUES('" + sku_no + "','" + vendor_no + "','" + vend_sku + "','" + last_price + "','" + qty + "','" + curr_key + "','" + m_manuf_lead_time + "','" + m_trasit_lead_time + "','" + m_quote_date + "','" + m_qrt + "','" + fob_at + "','" + fob_nl + "','" + fob_sg + "','" + ship_meth_at + "','" + ship_meth_nl + "','" + ship_meth_sg + "','" + m_qty_per_box + "')";
                  //  }
                  //  else
                 //   {
                 //       qry = qry + "   update dbo.vendor_sku  set sku_no='" + sku_no + "', vendor_no='" + vendor_no + "', vend_sku='" + vend_sku + "', last_price='" + last_price + "', qty='" + qty + "', curr_key='" + curr_key + "', m_manuf_lead_time='" + m_manuf_lead_time + "', m_trasit_lead_time='" + m_trasit_lead_time + "', m_quote_date='" + m_quote_date + "', m_qrt='" + m_qrt + "', fob_at='" + fob_at + "', fob_nl='" + fob_nl + "', fob_sg='" + fob_sg + "', ship_meth_at='" + ship_meth_at + "', ship_meth_nl='" + ship_meth_nl + "', ship_meth_sg='" + ship_meth_sg + "', m_qty_per_box='" + m_qty_per_box + "' where repl_id=" + rep_id.ToString();
                 //   }
                //}
            }

            if (!duplicatePart && !Qblank)
            {
                string constr = ConfigurationManager.ConnectionStrings["epConnectionString"].ConnectionString;
                qry = qryDelete + qry;                              //          +"  COMMIT ";
                qry = qry + "  exec   sp_update_part_price '" + vendor_no + "'";
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
                catch(Exception exp)
                {
                    string error = exp.Message.ToString();
                    tran.Rollback();
                    MessageBox.Show("Please check the data if part already exists for the quarter");
                }
                cn.Close();
                bindGrid();
            }
            
            else
            {
                if (Qblank)
                    MessageBox.Show("Quarter hasn't been given for one of the parts. Please check!");
                else if (duplicatePart)

                MessageBox.Show("Please check for duplicate parts");
                
            }
            Cursor.Current = Cursors.Default;
            
        }

        string filenameConfirm = "";
        string filenamePackaging = "";
        string filenamePackagingConfirm = "";
        
        
        private void btnExport_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
           
          //  filenameConfirm=CreateExcel("Vendor Quote Confirmation");
           // filenamePackaging = CreateExcel("Vendor Packaging Information");
            cellUpdate();
           // CreateExcel("Vendor Packaging Information");
            Cursor.Current = Cursors.Default;
        }

        DataSet dsCombo;
        public void loadComboData()
        {
            string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
            string sql = "SELECT [site],[fob],[ship_meth],[lead_time]  FROM [MIMDIST].[dbo].[m_pe_maintian_lt] ";
            dsCombo = getDataSet(sql, constr);          
        }

        DataSet dsComboCarrier;
        public void loadCarrier()
        {
            string constr = ConfigurationManager.ConnectionStrings["epConnectionString"].ConnectionString;
            string sql = " SELECT ship_via_code, ship_via_name FROM arshipv  ";
            dsComboCarrier = getDataSet(sql, constr);
            comboCarrier.DataSource = dsComboCarrier.Tables[0];
            comboCarrier.DisplayMember = "ship_via_name";
            comboCarrier.ValueMember = "ship_via_code";
            }


        public DataSet getDataSet(string sqlStr)
        {
            string constr = ConfigurationManager.ConnectionStrings["epConnectionString"].ConnectionString;
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlStr, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;
        }


        public string CreateExcel(string templatename,string filename)
        {

            // string  filename= this.vendor.Text + " " + this.ComboQtr.Text + " " + templatename + " _" + DateTime.Now.Year.ToString("00") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + ".xlsx";
            //  excelWorkbook.SaveAs("E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\"+filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            //string workbookPath = "E:\\Ruchi\\MaintainGoogleDoc\\MaintainGoogleDoc\\VendorSheet.xls";
           
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Vendor Quote";
            filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Vendor Quote\\" + filename;
            if ( ! System.IO.Directory.Exists(folder))
            {
                System.IO.Directory.CreateDirectory(folder);
            }
            
            if (System.IO.File.Exists(filename))
                System.IO.File.Delete(filename);


            //string workbookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\" +templatename+".xlsx";
            string workbookPath = "\\\\mvfile\\MV REPORTS\\Master Data & Forms\\PlanningExecutionForms\\VendorQuotes\\" + templatename + ".xlsx";
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
            return filename;

        }

        public int  sheetnumConfirm = 0;
        public int sheetnumPackaging = 0;
        public int sheetnumPackagingConfirm = 0;
        Microsoft.Office.Interop.Excel.Application excelAppConfirm;
        Microsoft.Office.Interop.Excel.Application excelAppPackaging;
        Microsoft.Office.Interop.Excel.Application excelAppPackagingConfirm;


        public void cellUpdate()
        {

            string part_num = "";
            string templateConf = "Vendor Quote Confirmation";
            string templatePackaging = "Vendor Packaging Information";
              string templatePackagingConf = "Vendor Quote and Pack Confirmation";

            foreach (UltraGridRow row in this.vendorGrid.Rows)
                if (row.Cells[row.Cells.Count - 1].Value.ToString() == "True")
                    part_num = part_num + "," + row.Cells[2].Value.ToString();
            if (part_num == "")
                MessageBox.Show("Select the parts you want to export");
            else
            {

                System.Diagnostics.Process p = new System.Diagnostics.Process();
                try
                {
                    // get data from dataset


                    String strusername = Environment.UserName;

                   // String strQuery = "select Quote_FileName,Pack_FileName, [Vendor Name],[Request Date],[Quote Quarter],[Customer Part Number],[MFG Part Number],[Last Quote Price]  from v_Vendor_Quote_Confirmation  ";
                    String strQuery = "SELECT TOP 1000 [VS_Quote_Qtr]      ,[VS_MIM_PN]      ,[VendCode]      ,[Pack_FileName]      ,[Vendor Name]      ,[Request Date]      ,[Quote Quarter]      ,[Customer Part Number]      ,[MFG Part Number]      ,[Last Quote Price]      ,[BoxQty]      ,[BoxL]      ,[BoxW]      ,[BoxH]      ,[BoxGW]      ,[PalletQty]      ,[PalletL]      ,[PalletW]      ,[PalletH]      ,[PalletGW]  FROM [MIMDIST].[dbo].[v_Vendor_QuotePack_Confirmation]";
                    string vendor_no = this.vendor.Text;
                    string qtr = this.cmbQtrFilter.Text;
                    if (vendor_no == "")
                        MessageBox.Show("Please select the vendor");
                    else
                    {
                        strQuery = strQuery + "  where   VendCode='" + vendor_no + "' and VS_MIM_PN in ('" + part_num.Substring(1, part_num.Length - 1).Replace(",", "','").ToString() + "')";
                    }
                    if (qtr != "")
                        strQuery = strQuery + " and [Quote Quarter]='" + qtr + "'    order by [Customer Part Number]";


/*                    String strQueryPackaging = "    select [vend_code]      ,[Vendor Name]      ,[Pack_FileName]      ,[Request Date]      ,[Quote Quarter]      ,[Customer Part Number]      ,[BoxQty]      ,[BoxL]      ,[BoxW]      ,[BoxH]      ,[BoxGW]      ,[PalletQty]      ,[PalletL]      ,[PalletW]      ,[PalletH]      ,[PalletGW]      ,[mim_pn]  from v_Vendor_Package_Confirmation  ";
                
                    if (vendor_no == "")
                        MessageBox.Show("Please select the vendor");
                    else
                    {
                        strQueryPackaging = strQueryPackaging + "  where   vend_code='" + vendor_no + "' and mim_pn in ('" + part_num.Substring(1, part_num.Length - 1).Replace(",", "','").ToString() + "')    ";
                    }
                    if (qtr != "")
                        strQueryPackaging = strQueryPackaging + " and [Quote Quarter]='" + qtr + "'    order by [Customer Part Number]    ";
*/


                    strQuery = strQuery + "   SELECT [first_name] + ' '+[last_name] as user_name ,[phone]  ,[email]   FROM [MIMDIST].[dbo].[m_mim_user_table] where user_name='" +strusername + "'";
                    string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
                    DataSet dsRFQ = getDataSet(strQuery, constr);
                    if  (dsRFQ.Tables.Count>0)
                        if (dsRFQ.Tables[0].Rows.Count > 0)
                        {

                            // filenameConfirm = dsRFQ.Tables[0].Rows[0]["Quote_FileName"].ToString() + ".xlsx";
                            //  filenamePackaging = dsRFQ.Tables[1].Rows[0]["Pack_FileName"].ToString() + ".xlsx";
                            filenamePackagingConfirm = dsRFQ.Tables[0].Rows[0]["Pack_FileName"].ToString() + ".xlsx";
                            //  filenameConfirm = CreateExcel(templateConf, filenameConfirm);
                            filenamePackagingConfirm = CreateExcel(templatePackagingConf, filenamePackagingConfirm);

                            
                            //**********************************************************************
                            //               Package and    Confirmation Excel
                            //*******************   ****************************************************

                            excelAppPackagingConfirm = new Microsoft.Office.Interop.Excel.Application();
                            string workbookPathPackagingConfirm = filenamePackagingConfirm;

                            Microsoft.Office.Interop.Excel.Workbook excelWorkbookPackagingConfirm = excelAppPackagingConfirm.Workbooks.Open(workbookPathPackagingConfirm, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                            Microsoft.Office.Interop.Excel.Sheets excelSheetsPackagingConfirm = excelWorkbookPackagingConfirm.Worksheets;
                            
                            sheetnumPackagingConfirm = 1;
                            excelAppPackagingConfirm.DisplayAlerts = false;
                            Microsoft.Office.Interop.Excel.Worksheet sheetPackagingConfirm = excelSheetsPackagingConfirm[1] as Microsoft.Office.Interop.Excel.Worksheet;
                            sheetPackagingConfirm.Name = this.cmbQtrFilter.Text;

                            int col = 0;
                            sheetPackagingConfirm.Cells[3, 10] = qtr;
                        
                            for (int i = 0; i < dsRFQ.Tables[0].Rows.Count; i++)
                            {
                                int row = i + 14;

                                // Vendor Name
                                sheetPackagingConfirm.Cells[2, 3] = dsRFQ.Tables[0].Rows[0]["Vendor Name"].ToString();
                                sheetPackagingConfirm.Cells[1, 19] = dsRFQ.Tables[0].Rows[0]["VendCode"].ToString();
                           
                            }

                            if (dsRFQ.Tables[1].Rows.Count > 0)
                            {
                                sheetPackagingConfirm.Cells[1, 14] = dsRFQ.Tables[1].Rows[0]["user_name"].ToString();
                                sheetPackagingConfirm.Cells[2, 14] = dsRFQ.Tables[1].Rows[0]["phone"].ToString();
                                sheetPackagingConfirm.Cells[3, 14] = dsRFQ.Tables[1].Rows[0]["email"].ToString();
                            }

                             for (int i = 0; i < dsRFQ.Tables[0].Rows.Count; i++)
                            {
                                int row = i + 15;
                                sheetPackagingConfirm.Cells[row, 1] = dsRFQ.Tables[0].Rows[i]["Customer Part Number"].ToString();
                                sheetPackagingConfirm.Cells[row, 2] = dsRFQ.Tables[0].Rows[i]["MFG Part Number"].ToString();
                                sheetPackagingConfirm.Cells[row, 3] = dsRFQ.Tables[0].Rows[i]["Last Quote Price"].ToString();
                               
                                sheetPackagingConfirm.Cells[row, 11] = dsRFQ.Tables[0].Rows[i]["BoxQty"].ToString();
                                sheetPackagingConfirm.Cells[row, 12] = dsRFQ.Tables[0].Rows[i]["BoxL"].ToString();
                                sheetPackagingConfirm.Cells[row, 13] = dsRFQ.Tables[0].Rows[i]["BoxW"].ToString();
                                sheetPackagingConfirm.Cells[row, 14] = dsRFQ.Tables[0].Rows[i]["BoxH"].ToString();
                                sheetPackagingConfirm.Cells[row, 15] = dsRFQ.Tables[0].Rows[i]["BoxGW"].ToString();
                                sheetPackagingConfirm.Cells[row, 16] = dsRFQ.Tables[0].Rows[i]["PalletQty"].ToString();
                                sheetPackagingConfirm.Cells[row, 17] = dsRFQ.Tables[0].Rows[i]["PalletL"].ToString();
                                sheetPackagingConfirm.Cells[row, 18] = dsRFQ.Tables[0].Rows[i]["PalletW"].ToString();
                                sheetPackagingConfirm.Cells[row, 19] = dsRFQ.Tables[0].Rows[i]["PalletH"].ToString();
                                sheetPackagingConfirm.Cells[row, 20] = dsRFQ.Tables[0].Rows[i]["PalletGW"].ToString();
                    
                            }
                        


                            excelWorkbookPackagingConfirm.Save();
                            excelWorkbookPackagingConfirm.Close(true, false, Type.Missing);

                            p.StartInfo.FileName = filenamePackagingConfirm;
                            //p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
                            p.Start();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppConfirm);






                            //***********************************************************************
                            //                  Confirmation Excel
                            //***********************************************************************

                            /*     excelAppConfirm = new Microsoft.Office.Interop.Excel.Application();

                                 string workbookPathConfirm = filenameConfirm;

                                 Microsoft.Office.Interop.Excel.Workbook excelWorkbookConfirm = excelAppConfirm.Workbooks.Open(workbookPathConfirm, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                                 Microsoft.Office.Interop.Excel.Sheets excelSheetsConfirm = excelWorkbookConfirm.Worksheets;


                                 sheetnumConfirm = 1;


                                 excelAppConfirm.DisplayAlerts = false;
                                 Microsoft.Office.Interop.Excel.Worksheet sheetConfirm = excelSheetsConfirm[1] as Microsoft.Office.Interop.Excel.Worksheet;
                                 sheetConfirm.Name = ComboQtr.Text;

                                 int col = 0;
                                 sheetConfirm.Cells[3, 6] = qtr;
                               //  sheetConfirm.Cells[1, 7] = dsRFQ.Tables[0].Rows[0]["Request Date"].ToString();
                                 for (int i = 0; i < dsRFQ.Tables[0].Rows.Count; i++)
                                 {
                                     int row = i + 14;

                                     // Vendor Name
                                     sheetConfirm.Cells[2, 3] = dsRFQ.Tables[0].Rows[0]["Vendor Name"].ToString();
                                     for (int j = 5; j < 8; j++)
                                     {
                                         col = j - 4;
                                         // excelApp.get_Range("A1:A360,B1:E1", Type.Missing).Merge(Type.Missing);
                                         sheetConfirm.Cells[row, col] = dsRFQ.Tables[0].Rows[i][j].ToString();
                                     }

                                 }

                                 if (dsRFQ.Tables[1].Rows.Count > 0)
                                 {
                                     sheetConfirm.Cells[1, 8] = dsRFQ.Tables[2].Rows[0][0].ToString();
                                     sheetConfirm.Cells[2, 8] = dsRFQ.Tables[2].Rows[0][1].ToString();
                                     sheetConfirm.Cells[3, 8] = dsRFQ.Tables[2].Rows[0][2].ToString();
                                 }

                                 excelWorkbookConfirm.Save();
                                 excelWorkbookConfirm.Close(true, false, Type.Missing);

                                 p.StartInfo.FileName = filenameConfirm;
                                 //p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
                                 p.Start();
                                 System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppConfirm);



                                 //***********************************************************************
                                 //                  Packaging Excel
                                 //***********************************************************************
                                 filenamePackaging = CreateExcel(templatePackaging, filenamePackaging);
                                 excelAppPackaging = new Microsoft.Office.Interop.Excel.Application();

                                 string workbookPathPackaging = filenamePackaging;
                                 Microsoft.Office.Interop.Excel.Workbook excelWorkbookPackaging = excelAppPackaging.Workbooks.Open(workbookPathPackaging, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                                 Microsoft.Office.Interop.Excel.Sheets excelSheetsPackaging = excelWorkbookPackaging.Worksheets;
                                 sheetnumPackaging = 1;


                                 excelAppPackaging.DisplayAlerts = false;
                                 Microsoft.Office.Interop.Excel.Worksheet sheetPackaging = excelSheetsPackaging[1] as Microsoft.Office.Interop.Excel.Worksheet;
                                 sheetPackaging.Name = ComboQtr.Text;

                                 col = 0;

                                 sheetPackaging.Cells[2, 3] = dsRFQ.Tables[1].Rows[0]["Vendor Name"].ToString();
                            //     sheetPackaging.Cells[1, 7] = dsRFQ.Tables[1].Rows[0]["Request Date"].ToString();
                                 for (int i = 0; i < dsRFQ.Tables[0].Rows.Count; i++)
                                 {
                                     int row = i + 17;
                                     sheetPackaging.Cells[row, 1] = dsRFQ.Tables[1].Rows[i]["Customer Part Number"].ToString();
                                     sheetPackaging.Cells[row, 2] = dsRFQ.Tables[1].Rows[i]["BoxQty"].ToString();
                                     sheetPackaging.Cells[row, 3] = dsRFQ.Tables[1].Rows[i]["BoxL"].ToString();
                                     sheetPackaging.Cells[row, 4] = dsRFQ.Tables[1].Rows[i]["BoxW"].ToString();
                                     sheetPackaging.Cells[row, 5] = dsRFQ.Tables[1].Rows[i]["BoxH"].ToString();
                                     sheetPackaging.Cells[row, 6] = dsRFQ.Tables[1].Rows[i]["BoxGW"].ToString();
                                     sheetPackaging.Cells[row, 7] = dsRFQ.Tables[1].Rows[i]["PalletQty"].ToString();
                                     sheetPackaging.Cells[row, 8] = dsRFQ.Tables[1].Rows[i]["PalletL"].ToString();
                                     sheetPackaging.Cells[row, 9] = dsRFQ.Tables[1].Rows[i]["PalletW"].ToString();
                                     sheetPackaging.Cells[row, 10] = dsRFQ.Tables[1].Rows[i]["PalletH"].ToString();
                                     sheetPackaging.Cells[row, 11] = dsRFQ.Tables[1].Rows[i]["PalletGW"].ToString();
                    
                                 }
                        


                                 excelWorkbookPackaging.Save();
                                 excelWorkbookPackaging.Close(true, false, Type.Missing);

                                 p.StartInfo.FileName = filenamePackaging;
                                 //p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
                                 p.Start();
                                 System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppPackaging);*/
                        }
                        else
                        {
                            MessageBox.Show("Part is not set up in Epicor ");
                        }
                    
                }
                catch (Exception e)
                {
                    string Message = e.Message;
                   // MessageBox.Show(Message);
                    if (excelAppConfirm != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppConfirm);
                }
                finally
                {
                    GC.Collect();
                    p.Close();
                }
            }


        }

        private void vendor_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
           
        }

        private void vendor_ListChanged(object sender, Infragistics.Win.ValueListChangedEventArgs e)
        {

        }
        private void vendor_RowSelected(object sender, RowSelectedEventArgs e)
        {
            string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
            this.ultraTextEditor2.Text = "";
            if (vendor.Text.Trim() != "")
            {
                string sql = "   SELECT distinct address_name FROM [MIMDIST].[dbo].[apmaster] where status_type=5 and [vend_class_code]='VENDOR' and  vendor_code='" + vendor.Text + "'";
                DataSet ds = getDataSet(sql, constr);
                if (ds.Tables.Count > 0)
                    if (ds.Tables[0].Rows.Count > 0)
                        this.ultraTextEditor2.Text = ds.Tables[0].Rows[0]["address_name"].ToString();
            }
        }


        private void ultraButton2_Click_1(object sender, EventArgs e)
        {
            

            
        }

        private void vendorGrid_AfterCellUpdate(object sender, CellEventArgs e)
        {
           // if (e.Cell.Column.Index ==21 )
            //{                
               // if (e.Cell.GetType() )
              //  vendorGrid.Rows[e.Cell.Row.Index].Selected = Convert.ToBoolean(e.Cell.Text);                
          //  }
          ///  else
         //   {
                Infragistics.Win.Appearance app1 = new Infragistics.Win.Appearance();
                app1.BackColor = Color.White;
                Infragistics.Win.Appearance app3 = new Infragistics.Win.Appearance();
                app3.BackColor = Color.Violet;
                Infragistics.Win.Appearance app2 = new Infragistics.Win.Appearance();
                app2.BackColor = Color.Red;
                string ship_meth = "";
                string fob = "";
                string site = "";
              //  MessageBox.Show(e.Cell.Column.Index.ToString());
                if (e.Cell.Column.Key =="Ship Meth AT")
                {
                    site = "AT";
                    ship_meth = e.Cell.Value.ToString();
                    fob = e.Cell.Row.Cells["FOB AT"].Value.ToString();
                }
                if (e.Cell.Column.Key == "FOB AT")
                {
                    site = "AT";
                    fob = e.Cell.Value.ToString();
                    ship_meth = e.Cell.Row.Cells["Ship Meth AT"].Value.ToString();
                }

                if (e.Cell.Column.Key == "Ship Meth NL")
                {
                    site = "NL";
                    ship_meth = e.Cell.Value.ToString();
                    fob = e.Cell.Row.Cells["FOB NL"].Value.ToString();
                }
                if (e.Cell.Column.Key == "FOB NL")
                {
                    site = "NL";
                    fob = e.Cell.Value.ToString();
                    ship_meth = e.Cell.Row.Cells["Ship Meth NL"].Value.ToString();
                }
        /*        if (e.Cell.Column.Index == 15)
                {
                    site = "SG";
                    ship_meth = e.Cell.Value.ToString();
                    fob = e.Cell.Row.Cells["FOB SG"].Value.ToString();
                }
                if (e.Cell.Column.Index ==12)
                {
                    site = "SG";
                    fob = e.Cell.Value.ToString();
                    ship_meth = e.Cell.Row.Cells["Ship Meth SG"].Value.ToString();
                }*/

                if (ship_meth != "" || fob != "")
                {
                    int leadTime = getATLeadTime(site, ship_meth, fob);
                    if (leadTime == 0)
                    {
                        //  e.Cell.Row.Appearance = app2;
                        // UltraGridCell cell = e.Cell as UltraGridCell;
                        // cell.Appearance = app2;
                        if (site == "AT")
                        {
                            (e.Cell.Row.Cells["FOB AT"] as UltraGridCell).Appearance = app2;
                            (e.Cell.Row.Cells["Ship Meth AT"] as UltraGridCell).Appearance = app2;
                        }
                        else if (site == "NL")
                        {
                            (e.Cell.Row.Cells["FOB NL"] as UltraGridCell).Appearance = app2;
                            (e.Cell.Row.Cells["Ship Meth NL"] as UltraGridCell).Appearance = app2;
                        }
              /*          else if (site == "SG")
                        {
                            (e.Cell.Row.Cells["FOB SG"] as UltraGridCell).Appearance = app2;
                            (e.Cell.Row.Cells["Ship Meth SG"] as UltraGridCell).Appearance = app2;
                        }*/
                    }

                    else
                    {
                        if (site == "AT")
                        {
                            (e.Cell.Row.Cells["FOB AT"] as UltraGridCell).Appearance = app1;
                            (e.Cell.Row.Cells["Ship Meth AT"] as UltraGridCell).Appearance = app1;
                        }
                        else if (site == "NL")
                        {
                            (e.Cell.Row.Cells["FOB NL"] as UltraGridCell).Appearance = app1;
                            (e.Cell.Row.Cells["Ship Meth NL"] as UltraGridCell).Appearance = app1;
                        }
             /*           else if (site == "SG")
                        {
                            (e.Cell.Row.Cells["FOB SG"] as UltraGridCell).Appearance = app1;
                            (e.Cell.Row.Cells["Ship Meth SG"] as UltraGridCell).Appearance = app1;
                        }*/
                    }
                    //e.Cell.Row.Appearance = app1;
                    if (site == "AT")
                        e.Cell.Row.Cells["AT Transit LT"].Value = leadTime.ToString();
                    else if (site == "NL")
                        e.Cell.Row.Cells["NL Transit LT"].Value = leadTime.ToString();
               /*     else if (site == "SG")
                        e.Cell.Row.Cells["SG Transit LT"].Value = leadTime.ToString();*/
                }
        //    }     

        }

        private void vendorGrid_CellChange(object sender, CellEventArgs e)
        {
           // MessageBox.Show(e.Cell.Column.Index.ToString());
         
        }

        public int getATLeadTime(string site,string ship_meth, string fob)
        {
            if (ship_meth != "" && fob != "")
            {
                DataRow[] dr = dsCombo.Tables[0].Select(" site='" + site + "' and ship_meth='" + ship_meth + "' and FOB='" + fob + "'");
                if (dr.Length > 0)
                   return  Convert.ToInt32(dr[0]["lead_time"].ToString());              
           }
            return 0;
        }
     

        private void vendorGrid_AfterCellListCloseUp(object sender, CellEventArgs e)
        {

        }

        private void comboCarrier_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Bands[0].Columns["ship_via_code"].Hidden = true;
        }

        private void ComboQtr_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Bands[0].Columns["ord"].Hidden = true;
        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }
                  
        }

  
}
