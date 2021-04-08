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
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;

namespace Version3
{
    public partial class PartMaintenance : Form
    {
        string constrEpicor = ConfigurationManager.ConnectionStrings["epConnectionString"].ConnectionString;
        string constrMVERP = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        DataSet dsALL;

        private void ultraGrid1_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Override.AllowMultiCellOperations = AllowMultiCellOperation.All;
            e.Layout.Override.CellClickAction = CellClickAction.CellSelect;
            this.ultraGrid1.DisplayLayout.Override.AllowAddNew = AllowAddNew.Yes;
            this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = SelectType.Extended;
        }

        public PartMaintenance()
        {           
            InitializeComponent();
            dsALL = getDataSet();
            loadPartType(dsALL.Tables[0]);
            loadManufacturer(dsALL.Tables[0]);
            loadComboData();
            loadEditFOB(dsCombo.Tables[0]);  
        }

        DataSet dsCombo;
        public void loadComboData()
        {
            string sql = "SELECT [site],[fob],[ship_meth],[lead_time]  FROM [MIMDIST].[dbo].[m_pe_maintian_lt]";
            dsCombo = getDataSet(sql);
        }

        public DataSet getDataSet(string sqlStr)
        {
            SqlConnection cn = new SqlConnection(constrMVERP);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlStr, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;
        }

        private void loadSiteCode(DataTable dt)
        {
            DataTable dtSiteCode = dt.DefaultView.ToTable(true, "site");
            DataView dv = dtSiteCode.DefaultView;
            dv.Sort = "site";
            dtSiteCode = dv.ToTable();         
        }


        private void loadEditFOB(DataTable dt)
        {
            DataTable dtATFOB = dt.Select("site='AT'").CopyToDataTable().DefaultView.ToTable(true, "fob");
            DataTable dtNLFOB = dt.Select("site='NL'").CopyToDataTable().DefaultView.ToTable(true, "fob");
            DataTable dtSGFOB = dt.Select("site='SG'").CopyToDataTable().DefaultView.ToTable(true, "fob");     
          
            comboATFOB.DataSource = dtATFOB;
            comboATFOB.DisplayMember = "fob";
            comboATFOB.ValueMember = "fob";

            cmbATFOB.DataSource = dtATFOB;
            cmbATFOB.DisplayMember = "fob";
            cmbATFOB.ValueMember = "fob";


            comboNLFOB.DataSource = dtNLFOB;
            comboNLFOB.DisplayMember = "fob";
            comboNLFOB.ValueMember = "fob";


            cmbNLFOB.DataSource = dtNLFOB;
            cmbNLFOB.DisplayMember = "fob";
            cmbNLFOB.ValueMember = "fob";


            comboSGFOB.DataSource = dtSGFOB;
            comboSGFOB.DisplayMember = "fob";
            comboSGFOB.ValueMember = "fob";

            cmbSGFOB.DataSource = dtSGFOB;
            cmbSGFOB.DisplayMember = "fob";
            cmbSGFOB.ValueMember = "fob";
            //comboATFOB.SelectedText = "ALL";
          //  comboATFOB.SelectedText = "ALL";

        }

        private void loadATEditShipMethod(DataTable dt)
        {
            if (this.comboATFOB.SelectedText.ToString()!="")
            {
                DataTable dtShipMeth = dt.Select("site='AT' and FOB='"+this.comboATFOB.SelectedText.ToString()+"'").CopyToDataTable().DefaultView.ToTable(true, "ship_meth");
                DataView dv = dtShipMeth.DefaultView;
                dv.Sort = "ship_meth";
                dtShipMeth = dv.ToTable();
            
                //   DataRow drShipMeth = dtShipMeth.NewRow();
                //   drShipMeth["ship_meth"] = "ALL";
                //   dtShipMeth.Rows.InsertAt(drShipMeth, 0);

                this.comboATShipMeth.DataSource = dtShipMeth;
                comboATShipMeth.DisplayMember = "ship_meth";
                comboATShipMeth.ValueMember = "ship_meth";      
            }
          /*  this.comboNLShipMeth.DataSource = dtShipMeth;
            comboNLShipMeth.DisplayMember = "ship_meth";
            comboNLShipMeth.ValueMember = "ship_meth";
       
            this.comboATShipMeth.DataSource = dtShipMeth;
            comboSGShipMeth.DisplayMember = "ship_meth";
            comboSGShipMeth.ValueMember = "ship_meth";*/
          }

        private void loadPartType(DataTable dt)
        {
            DataTable dtPartType = dt.DefaultView.ToTable(true, "Commodity");
            DataView dv = dtPartType.DefaultView;
            dv.Sort = "Commodity";
            dtPartType = dv.ToTable();
            DataRow drPartType = dtPartType.NewRow();
            drPartType["Commodity"] = "ALL";
            dtPartType.Rows.InsertAt(drPartType, 0);
            cmbPartType.DataSource = dtPartType;
            cmbPartType.DisplayMember = "Commodity";
            cmbPartType.ValueMember = "Commodity";
        }
        
      
        private void loadManufacturer(DataTable dt)
        {
            DataTable dtManufacturer = dt.DefaultView.ToTable(true, "Manufacturer");
            DataView dv = dtManufacturer.DefaultView;
            dv.Sort = "Manufacturer";
            dv.RowFilter = "Manufacturer <>''"; 
            dtManufacturer = dv.ToTable();
            DataRow drManf = dtManufacturer.NewRow();
            drManf["Manufacturer"] = "ALL";
            dtManufacturer.Rows.InsertAt(drManf, 0);
            cmbManufacturer.DataSource = dtManufacturer;
            cmbManufacturer.DisplayMember = "Manufacturer";
            cmbManufacturer.ValueMember = "Manufacturer";
        }


        private void loadATFOB(DataTable dt)
        {
            DataTable dtATFOB = dt.DefaultView.ToTable(true, "AT FOB");
            DataView dv = dtATFOB.DefaultView;
            dv.Sort = "AT FOB";

            dv.RowFilter = "[AT FOB] <>''";
            dtATFOB = dv.ToTable();
            
            DataRow dr = dtATFOB.NewRow();
            dr["AT FOB"] = "ALL";
            
            //dtATFOB.Rows.InsertAt(dr, 0);
            cmbATFOB.DataSource = dtATFOB;
            cmbATFOB.DisplayMember = "AT FOB";
            cmbATFOB.ValueMember = "AT FOB";


        }
        private void loadEUFOB(DataTable dt)
        {
            DataTable dtEUFOB = dt.DefaultView.ToTable(true, "NL FOB");

            DataView dv = dtEUFOB.DefaultView;
            dv.Sort = "NL FOB";
            dv.RowFilter = "[NL FOB] <>''";
            DataRow dr = dtEUFOB.NewRow();

            dr["NL FOB"] = "ALL";
          //  dtEUFOB.Rows.InsertAt(dr, 0);
            dtEUFOB = dv.ToTable();
            cmbNLFOB.DataSource = dtEUFOB;
            cmbNLFOB.DisplayMember = "NL FOB";
            cmbNLFOB.ValueMember = "NL FOB";
        }
        
        private void loadSGFOB(DataTable dt)
        {
            DataTable dtSGFOB = dt.DefaultView.ToTable(true, "SG FOB");
            DataView dv = dtSGFOB.DefaultView;
            dv.Sort = "SG FOB";

            dv.RowFilter = "[SG FOB] <>''";
            dtSGFOB = dv.ToTable();

            DataRow dr = dtSGFOB.NewRow();
            dr["SG FOB"] = "ALL";

        //    dtSGFOB.Rows.InsertAt(dr, 0);
            cmbSGFOB.DataSource = dtSGFOB;
            cmbSGFOB.DisplayMember = "AT FOB";
            cmbSGFOB.ValueMember = "AT FOB";


        }
        private void loadATShipMeth(DataTable dt)
        {
            DataTable dtATShipMeth = dt.DefaultView.ToTable(true, "AT SHIP METHOD");

            DataView dv = dtATShipMeth.DefaultView;
            dv.Sort = "AT SHIP METHOD";
            dv.RowFilter = "[AT SHIP METHOD] <>''";
            DataRow dr = dtATShipMeth.NewRow();
            dr["AT SHIP METHOD"] = "ALL";
          //  dtATShipMeth.Rows.InsertAt(dr, 0);
            dtATShipMeth = dv.ToTable();
            cmbATShipMeth.DataSource = dtATShipMeth;
            cmbATShipMeth.DisplayMember = "AT SHIP METHOD";
            cmbATShipMeth.ValueMember = "AT SHIP METHOD";
        }
     
        private void loadEUShipMeth(DataTable dt)
        {
            DataTable dtNLShipMeth = dt.DefaultView.ToTable(true, "NL SHIP METHOD");
            DataView dv = dtNLShipMeth.DefaultView;
            dv.Sort = "NL SHIP METHOD";
            dv.RowFilter = "[NL SHIP METHOD] <>''";
            DataRow dr = dtNLShipMeth.NewRow();
            dr["NL SHIP METHOD"] = "ALL";
           // dtNLShipMeth.Rows.InsertAt(dr, 0);
            dtNLShipMeth = dv.ToTable();
            cmbNLShipMeth.DataSource = dtNLShipMeth;
            cmbNLShipMeth.DisplayMember = "NL SHIP METHOD";
            cmbNLShipMeth.ValueMember = "NL SHIP METHOD";
        }

        //private void loadCategory(DataTable dt)
        //{
        //    DataTable dtCategory = dt.DefaultView.ToTable(true, "category_code", "Category");
        //    DataView dvCat = dtCategory.DefaultView;
        //    dvCat.Sort = "category_code";
        //    dtCategory = dvCat.ToTable();
        //    DataRow dr=dtCategory.NewRow();
        //    dr["category_code"] = 0;
        //    dr["Category"] = "ALL";
        //    dtCategory.Rows.InsertAt(dr,0);
        //    //  dtPartType.Select("type_code=" + cmbPartType.SelectedText);
        //   // cmbCategory.DataSource = dtCategory;
        //  //  cmbCategory.DisplayMember = "Category";
        // //   cmbCategory.ValueMember = "category_code";
           
        // // cmbCategory.Items.Insert(0,obj);
            
        //}
        DataTable finalDt;
        public void loadGrid()
        {
            gridviewPartMaintain.Columns.Clear();
            gridviewPartMaintain.DataSource = null;
            dsALL = getDataSet();
            DataView dvcatGrid = dsALL.Tables[0].DefaultView;

            string filterExp = "";
            if (this.txtPartNumber.Text != "" && this.txtPartNumber.Text != "ALL")
            {

                filterExp += "[Part Num] like '%" + this.txtPartNumber.Text + "%'";
            }
            if (this.cmbManufacturer.Text != "" && this.cmbManufacturer.Text != "ALL")
            {
                if (filterExp != "")
                    filterExp += " and ";
                filterExp += "[Manufacturer] = '" + this.cmbManufacturer.Text + "'";
            }
            if (this.cmbPartType.Text != "" && this.cmbPartType.Text != "ALL")
            {
                if (filterExp != "")
                    filterExp += " and ";
                filterExp += "[Commodity] = '" + this.cmbPartType.Text + "'";
            }
            //if (cmbCategory.Text != "ALL")
            //{
            //    if (filterExp != "")
            //        filterExp += " and ";

            //    filterExp += "Category='" + cmbCategory.Text + "'";
            //}
            if (this.cmbATFOB.Text != "")
            {
                if (filterExp != "")
                    filterExp += " and ";
                filterExp += "[AT FOB]='" + cmbATFOB.Text + "'";
            }
            if (cmbNLFOB.Text != "")
            {
                if (filterExp != "")
                    filterExp += " and ";
                filterExp += "[NL FOB]='" + cmbNLFOB.Text + "'";
            }


            if (this.cmbSGFOB.SelectedText != "" && this.cmbSGFOB.Text != "ALL")
            {
                if (filterExp != "")
                    filterExp += " and ";
                filterExp += "[SG FOB]='" + this.cmbSGFOB.SelectedText + "'";
            }

            if (this.cmbATShipMeth.Text != "" && this.cmbATShipMeth.Text != "ALL")
            {
                if (filterExp != "")
                    filterExp += " and ";
                filterExp += "[AT Ship Meth]='" + cmbATShipMeth.Text + "'";
            }

            if (this.cmbNLShipMeth.Text != "" && this.cmbNLShipMeth.Text != "ALL")
            {
                if (filterExp != "")
                    filterExp += " and ";

                filterExp += "[NL Ship Meth]='" + cmbNLShipMeth.Text + "'";
            }
            if (this.cmbSGShipMeth.Text != "" && this.cmbSGShipMeth.Text != "ALL")
            {
                if (filterExp != "")
                    filterExp += " and ";
                filterExp += "[SG Ship Meth]='" + cmbSGShipMeth.Text + "'";
            }
          
            dvcatGrid.RowFilter = filterExp;
            dvcatGrid.Sort = "  [Part Num] Asc";
            finalDt = dvcatGrid.ToTable();
            gridviewPartMaintain.DataSource = dvcatGrid.ToTable();
            ultraGrid1.DataSource = dvcatGrid.ToTable();   
       
            int cnt = 0;
            for (cnt = 0; cnt < gridviewPartMaintain.Columns.Count; cnt++)
            {
                if ((cnt != 5))  //(cnt != 13) && (cnt != 14) && (cnt != 8) &&
                    gridviewPartMaintain.Columns[cnt].DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
            }

            gridviewPartMaintain.AutoResizeColumns();
            gridviewPartMaintain.Columns[0].Visible = true;
            gridviewPartMaintain.Columns[1].Visible = true;
            gridviewPartMaintain.Columns[2].Visible = true;
            gridviewPartMaintain.Columns[2].ReadOnly =true;

            gridviewPartMaintain.Columns[0].ReadOnly = true;
            this.gridviewPartMaintain.Columns[0].Width = 250;
            gridviewPartMaintain.Columns[1].ReadOnly = true;
            this.gridviewPartMaintain.Columns[1].Width = 150;
            gridviewPartMaintain.Columns[2].ReadOnly = true;
            this.gridviewPartMaintain.Columns[2].Width =80;

            gridviewPartMaintain.Columns[3].ReadOnly = true;
            this.gridviewPartMaintain.Columns[3].Width = 100;           
         
            this.gridviewPartMaintain.Columns[3].Width = 100;
            
            DataGridViewComboBoxColumn cmbATShipMethod = new DataGridViewComboBoxColumn();
            cmbATShipMethod.HeaderText = "AT Ship Meth";
            cmbATShipMethod.Name = "ATcmb";
            cmbATShipMethod.Items.Add("GROUND");
            cmbATShipMethod.Items.Add("STANDARD AIR");
            cmbATShipMethod.Items.Add("OCEAN");
            
            DataGridViewComboBoxColumn cmbNLShipMethod = new DataGridViewComboBoxColumn();
            cmbNLShipMethod.HeaderText = "NL Ship Meth";
            cmbNLShipMethod.Name = "NLcmb";
            cmbNLShipMethod.Items.Add("GROUND");
            cmbNLShipMethod.Items.Add("STANDARD AIR");
            cmbNLShipMethod.Items.Add("OCEAN");
            
            DataGridViewComboBoxColumn cmbSGShipMethod = new DataGridViewComboBoxColumn();
            cmbSGShipMethod.HeaderText = "SG Ship Meth";
            cmbSGShipMethod.Name = "SGcmb";
            cmbSGShipMethod.Items.Add("GROUND");
            cmbSGShipMethod.Items.Add("STANDARD AIR");
            cmbSGShipMethod.Items.Add("OCEAN");

            DataGridViewComboBoxColumn ATcmbFOB = new DataGridViewComboBoxColumn();
            ATcmbFOB.HeaderText = "AT FOB";
            ATcmbFOB.Name = "ATcmbFOB";
            ATcmbFOB.Items.Add("ASIA");
            ATcmbFOB.Items.Add("EUROPE");
            ATcmbFOB.Items.Add("US");

            DataGridViewComboBoxColumn NLcmbFOB = new DataGridViewComboBoxColumn();
            NLcmbFOB.HeaderText = "NL FOB";
            NLcmbFOB.Name = "NLcmbFOB";
            NLcmbFOB.Items.Add("ASIA");
            NLcmbFOB.Items.Add("EUROPE");
            NLcmbFOB.Items.Add("US");

            DataGridViewComboBoxColumn cmbSGFOB = new DataGridViewComboBoxColumn();
            cmbSGFOB.HeaderText = "SG FOB";
            cmbSGFOB.Name = "SGcmbFOB";
            cmbSGFOB.Items.Add("ASIA");
            cmbSGFOB.Items.Add("EUROPE");
            cmbSGFOB.Items.Add("US");
       /*   gridviewPartMaintain.Columns.Insert(6,ATcmbFOB);
            gridviewPartMaintain.Columns.Insert(7, cmbATShipMethod);
            gridviewPartMaintain.Columns.Insert(9, NLcmbFOB);
            gridviewPartMaintain.Columns.Insert(10, cmbNLShipMethod);
            gridviewPartMaintain.Columns.Insert(12, cmbSGFOB);
            gridviewPartMaintain.Columns.Insert(13, cmbSGShipMethod);
*/
            gridviewPartMaintain.Columns[4].ReadOnly =false;
            this.gridviewPartMaintain.Columns[4].Width =450;

            gridviewPartMaintain.Columns[5].ReadOnly = false;
            this.gridviewPartMaintain.Columns[5].Width = 80;
           
          //  gridviewPartMaintain.Columns[6].ReadOnly = true;
          //  this.gridviewPartMaintain.Columns[6].Width =50;
            
            gridviewPartMaintain.Columns[7].ReadOnly = true;
            this.gridviewPartMaintain.Columns[7].Width = 80;
            gridviewPartMaintain.Columns[8].ReadOnly = true;
            this.gridviewPartMaintain.Columns[8].Width = 80;
            gridviewPartMaintain.Columns[9].ReadOnly = true;
            gridviewPartMaintain.Columns[10].ReadOnly = true;
            this.gridviewPartMaintain.Columns[9].Width = 50;
            this.gridviewPartMaintain.Columns[10].Width = 80;
            gridviewPartMaintain.Columns[11].ReadOnly = true;
            this.gridviewPartMaintain.Columns[11].Width = 80;
            gridviewPartMaintain.Columns[12].ReadOnly = true;
            this.gridviewPartMaintain.Columns[12].Width = 50;
            gridviewPartMaintain.Columns[13].ReadOnly = true;
            this.gridviewPartMaintain.Columns[13].Width = 80;
            gridviewPartMaintain.Columns[14].ReadOnly = true;
            this.gridviewPartMaintain.Columns[14].Width = 80;
         

             gridviewPartMaintain.EditingControlShowing +=new DataGridViewEditingControlShowingEventHandler(gridviewPartMaintain_EditingControlShowing);
            //gridviewPartMaintain.Columns[10].ReadOnly = true;
            //this.gridviewPartMaintain.Columns[10].Width = 100;

            
            //gridviewPartMaintain.Columns[11].ReadOnly = true;
            //this.gridviewPartMaintain.Columns[11].Width = 100;

            //gridviewPartMaintain.Columns[12].ReadOnly = true;
            //this.gridviewPartMaintain.Columns[12].Width = 100;

            //this.gridviewPartMaintain.Columns[14].Width = 600;
            //gridviewPartMaintain.Columns[15].ReadOnly = true;
            //this.gridviewPartMaintain.Columns[15].Width = 300;


            //gridviewPartMaintain.Columns[16].ReadOnly = true;
            //this.gridviewPartMaintain.Columns[16].Width = 300;

            //gridviewPartMaintain.Columns["AT Factory Lead Time" ].Width = 70;
            //gridviewPartMaintain.Columns["AT Transit Lead Time"].Width = 70;
            //gridviewPartMaintain.Columns["AT Total Lead Time"].Width = 70;
            //gridviewPartMaintain.Columns["EU Factory Lead Time"].Width = 70;
            //gridviewPartMaintain.Columns["EU Transit Lead Time"].Width = 70;
            //gridviewPartMaintain.Columns["EU Total Lead Time"].Width = 70; 
           
            //this.gridviewPartMaintain.Columns["AT Factory Lead Time"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //this.gridviewPartMaintain.Columns["AT Transit Lead Time"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //this.gridviewPartMaintain.Columns["AT Total Lead Time"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //this.gridviewPartMaintain.Columns["EU Factory Lead Time"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //this.gridviewPartMaintain.Columns["EU Transit Lead Time"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //this.gridviewPartMaintain.Columns["EU Total Lead Time"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; 

           // gridviewPartMaintain.Columns[15].ReadOnly = true;         
             

        }


        
  
       
private void gridviewPartMaintain_EditingControlShowing(object sender,DataGridViewEditingControlShowingEventArgs e)
{
            ComboBox dgvCombo = e.Control as ComboBox;
            if (dgvCombo != null)
            {
                // first remove event handler to keep from attaching multiple:
                dgvCombo.SelectedIndexChanged -= new  EventHandler(dvgCombo_SelectedIndexChanged);
                // now attach the event handler                
                dgvCombo.SelectedIndexChanged += new EventHandler(dvgCombo_SelectedIndexChanged);
            }

  //  ((ComboBox)e.Control).SelectionChangeCommitted += new EventHandler(ComboBox_SelectionChangeCommitted);
}
private void dvgCombo_SelectedIndexChanged(object sender, EventArgs e)
{
    ComboBox cmb = (ComboBox)sender;
    for (int i = 0; i < cmb.Items.Count; i++)
    {
        //List<string> lst=(List<string> )cmb.Items;
        string val = (string)cmb.Items[i];
        if (val== cmb.SelectedItem.ToString())
        {
            cmb.SelectedIndex = i;
            break;
        }
    }
 
    
}

        public DataSet getDataSet()
        {
            //string strqry = "select commodity as [Commodity],manufacturer as [Manufacturer],part_num as [Part Num],part_description as [Part Desc],factory_lead_time as [Factory Lead time],fob_at as [AT FOB],ship_meth_at as [AT Ship Method],transit_lt_at as [AT Transit Lead Time ],fob_NL as [NL FOB],ship_meth_nl as [NL Ship Method],transit_lt_nl as [NL Transit Lead Time ],fob_sg as [SG FOB],ship_meth_sg as [ SG Ship Method],transit_lt_sg as [SG Transit Lead Time ] from m_pe_maintain_lead_time";//select type_code,category_code,part_type_desc as[Part Type],Category_desc as [Category], abc_code as [ABC Code],part_num as [Part Num],at_fob_point as [AT FOB],at_ship_meth as [AT Ship Method],cast(m_at_factory_lt as integer) as [AT Factory Lead Time],cast(m_at_transit_lt as integer) as [AT Transit Lead Time],cast(m_at_factory_lt as integer) + cast(m_at_transit_lt as integer)  as [AT Total Lead Time],eu_fob_point as [EU FOB],eu_ship_meth as [EU Ship Method],cast(m_eu_factory_lt as integer) as [EU Factory Lead Time],cast(m_eu_tranit_lt as integer) as [EU Transit Lead Time],cast((m_eu_factory_lt+ m_eu_,Description,Manufacturer,Man_part_num as [Manf Part Num] from m_pe_maintain_lead_time  order by part_num";
            string strqry = " select commodity as [Commodity],manufacturer as [Manufacturer],supplytype AS Owner,partnumber as [Part Num],partdescription as [Part Desc],cast(FactorLT as integer) as [Fact LT],FOB_AT as [AT FOB],shipmeth_at  as [AT Ship Meth],cast(tranlt_at as integer) as [AT Tran LT],fob_NL  as [NL FOB],shipmeth_nl  as [NL Ship Meth],cast(tranlt_nl as integer) as [NL Tran LT],fob_sg  as [SG FOB] ,shipmeth_sg  as [SG Ship Meth],cast(tranlt_sg as integer) as [SG Tran LT] ,qty_per_box as [Box QTY],boxes_per_pallet as [Pallet QTY],ord_mult as [Ord Mult],LockedForEdit    from pe_app_part_maintenance order by partnumber ";
            SqlConnection cn = new SqlConnection(constrEpicor);
            cn.Open();
            SqlCommand cmd = new SqlCommand(strqry, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;
        }

        private void cmbPartType_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataView dvcatPart = dsALL.Tables[0].DefaultView;
            if (cmbPartType.Text.ToString() != "System.Data.DataRowView")
            {
                dvcatPart.RowFilter = "[commodity]='" + cmbPartType.Text + "'";
            }
           // loadCategory(dvcatPart.ToTable());
        }

        private void cmbCategory_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnDownloadGrid_Click(object sender, EventArgs e)
        {          
              loadGrid();           
            

        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
             string updateQry="";
             DataTable dtSave = gridviewPartMaintain.DataSource as DataTable;
              for (int i=0;i<dtSave.Rows.Count;i++)
                {
                    updateQry += "    update epicor.mimdist.dbo.part_price set " + " m_factory_lt ='" + dtSave.Rows[i]["Factory Lead Time"].ToString() + "'   where part_no='" + dtSave.Rows[i]["Part Num"].ToString() + "'";
                }

            //  MessageBox.Show(updateQry); 
              SqlConnection cn = new SqlConnection(constrEpicor);
          cn.Open();
          SqlCommand cmd = new SqlCommand(updateQry, cn);
          int roweffected=cmd.ExecuteNonQuery();
          if (roweffected > 0)
          {
              MessageBox.Show(roweffected + " rows are saved successfully ");
          }
          cn.Close();
          dsALL = getDataSet();         
          loadGrid();
          // return ds;

        }
        void gridviewPartMaintain_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // Clear the row error in case the user presses ESC.  
            //if ((e.ColumnIndex==8) || (e.ColumnIndex==9))
            //{
            //    int t1 = Convert.ToInt32(gridviewPartMaintain.Rows[e.RowIndex].Cells[8].Value);
            //    int t2 = Convert.ToInt32(gridviewPartMaintain.Rows[e.RowIndex].Cells[9].Value);
            //    gridviewPartMaintain.Rows[e.RowIndex].Cells[10].Value = (t1 + t2).ToString();
            //}
            //if ((e.ColumnIndex == 13) || (e.ColumnIndex == 14))
            //{
            //    int t1 = Convert.ToInt32(gridviewPartMaintain.Rows[e.RowIndex].Cells[13].Value);
            //    int t2 = Convert.ToInt32(gridviewPartMaintain.Rows[e.RowIndex].Cells[14].Value);
            //    gridviewPartMaintain.Rows[e.RowIndex].Cells[15].Value = (t1 + t2).ToString();
            //}

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {

        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void gridviewPartMaintain_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        
        }


        void gridviewPartMaintain_CurrentCellDirtyStateChanged(object sender,    EventArgs e)
        {
            if (gridviewPartMaintain.IsCurrentCellDirty)
            {
                gridviewPartMaintain.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void gridviewPartMaintain_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        { 
            /* if (gridviewPartMaintain.Columns[e.ColumnIndex].HeaderText == "AT FOB")
                {

                    DataGridViewCell cmb = (DataGridViewCell)gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    if (finalDt.Rows[e.RowIndex][9].ToString() != "")
                           cmb.Value = finalDt.Rows[e.RowIndex][9].ToString();
                    if (gridviewPartMaintain.Rows[e.RowIndex].Cells[2].Value.ToString() == "CONSIGNED")
                    {
                        if (finalDt.Rows[e.RowIndex][9].ToString() == "")
                            gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = false;
                    }
                    else
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = true;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.LightBlue;
                    }
                }

                if (gridviewPartMaintain.Columns[e.ColumnIndex].HeaderText == "NL FOB")
                {

                    DataGridViewCell cmb = (DataGridViewCell)gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    if (finalDt.Rows[e.RowIndex][11].ToString() != "")
                    cmb.Value = finalDt.Rows[e.RowIndex][11].ToString();
                    if (gridviewPartMaintain.Rows[e.RowIndex].Cells[2].Value.ToString() == "CONSIGNED")
                    {
                        if (finalDt.Rows[e.RowIndex][11].ToString() == "")
                            gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = false;
                    }
                    else
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = true;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.LightBlue;
                    }
                }

                if (gridviewPartMaintain.Columns[e.ColumnIndex].HeaderText == "SG FOB")
                {
                    DataGridViewCell cmb = (DataGridViewCell)gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    if (finalDt.Rows[e.RowIndex][13].ToString() != "")
                    cmb.Value = finalDt.Rows[e.RowIndex][13].ToString();
                    if (gridviewPartMaintain.Rows[e.RowIndex].Cells[2].Value.ToString() == "CONSIGNED")
                    {
                        if (finalDt.Rows[e.RowIndex][13].ToString() == "")
                            gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = false;
                    }
                    else
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = true;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.LightBlue;
                    }
                }


                if (gridviewPartMaintain.Columns[e.ColumnIndex].HeaderText == "AT Ship Meth")
                {

                    DataGridViewCell cmb = (DataGridViewCell)gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    if (finalDt.Rows[e.RowIndex][10].ToString() != "")
                    cmb.Value = finalDt.Rows[e.RowIndex][10].ToString();
                    if (gridviewPartMaintain.Rows[e.RowIndex].Cells[2].Value.ToString() == "CONSIGNED")
                    {
                        if (finalDt.Rows[e.RowIndex][10].ToString() == "")
                            gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = false;
                    }
                    else
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = true;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.LightBlue;
                    }
                }

                if (gridviewPartMaintain.Columns[e.ColumnIndex].HeaderText == "NL Ship Meth")
                {
                    DataGridViewCell cmb = (DataGridViewCell)gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    if (finalDt.Rows[e.RowIndex][12].ToString()!="")
                         cmb.Value = finalDt.Rows[e.RowIndex][12].ToString();
                    if (gridviewPartMaintain.Rows[e.RowIndex].Cells[2].Value.ToString() == "CONSIGNED")
                    {
                        if (finalDt.Rows[e.RowIndex][12].ToString() == "")
                            gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = false;
                    }
                    else
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = true;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.LightBlue;
                    }
                }

                if (gridviewPartMaintain.Columns[e.ColumnIndex].HeaderText == "SG Ship Meth")
                {

                    DataGridViewCell cmb = (DataGridViewCell)gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    if (finalDt.Rows[e.RowIndex][14].ToString() != "")
                             cmb.Value = finalDt.Rows[e.RowIndex][14].ToString();
                    if (gridviewPartMaintain.Rows[e.RowIndex].Cells[2].Value.ToString() == "CONSIGNED")
                    {
                        if (finalDt.Rows[e.RowIndex][14].ToString() == "")
                            gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = false;
                    }
                    else
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = true;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.LightBlue;
                    }
                }
                if (gridviewPartMaintain.Columns[e.ColumnIndex].HeaderText == "AT Transit Lead Time")
                {
                    if (gridviewPartMaintain.Rows[e.RowIndex].Cells[2].Value.ToString() != "CONSIGNED")
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = true;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.LightBlue;
                    }
                    else
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = false;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.White;
                    }
                }
                if (gridviewPartMaintain.Columns[e.ColumnIndex].HeaderText == "NL Transit Lead Time")
                {
                    if (gridviewPartMaintain.Rows[e.RowIndex].Cells[2].Value.ToString() != "CONSIGNED")
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = true;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.LightBlue;
                    }
                    else
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = false;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.White;
                    }
                }
                if (gridviewPartMaintain.Columns[e.ColumnIndex].HeaderText == "SG Transit Lead Time")
                {
                    if (gridviewPartMaintain.Rows[e.RowIndex].Cells[2].Value.ToString() != "CONSIGNED")
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = true;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.LightBlue;
                    }
                    else
                    {
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = false;
                        gridviewPartMaintain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = System.Drawing.Color.White;
                    }
                }*/

            }

        private void cmbSiteCode_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {

        }

        private void gridviewPartMaintain_SelectionChanged(object sender, EventArgs e)
        {
            if (gridviewPartMaintain.SelectedRows.Count > 0)
            {
                
                comboATFOB.Text = gridviewPartMaintain.SelectedRows[0].Cells[6].Value.ToString();
                comboATShipMeth.Text = gridviewPartMaintain.SelectedRows[0].Cells[7].Value.ToString();
                this.txtATLeadTime.Text = gridviewPartMaintain.SelectedRows[0].Cells[8].Value.ToString();
                comboNLFOB.Text = gridviewPartMaintain.SelectedRows[0].Cells[9].Value.ToString();
                comboNLShipMeth.Text = gridviewPartMaintain.SelectedRows[0].Cells[10].Value.ToString();
                this.txtNLLeadTime.Text = gridviewPartMaintain.SelectedRows[0].Cells[11].Value.ToString();
                comboSGFOB.Text = gridviewPartMaintain.SelectedRows[0].Cells[12].Value.ToString();
                comboSGShipMeth.Text = gridviewPartMaintain.SelectedRows[0].Cells[13].Value.ToString();
                this.txtSGLeadTime.Text = gridviewPartMaintain.SelectedRows[0].Cells[14].Value.ToString();
                if (gridviewPartMaintain.SelectedRows[0].Cells[1].Value.ToString() == "Y")
                {
                    comboATFOB.Enabled = false;
                    comboATShipMeth.Enabled = false;
                    txtATLeadTime.Enabled = false;
                    comboNLFOB.Enabled = false;
                    comboNLShipMeth.Enabled = false;
                    txtNLLeadTime.Enabled = false;
                    comboSGFOB.Enabled = false;
                }
            }
        }

        private void PartMaintenance_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dataSet3.vendor_sku' table. You can move, or remove it, as needed.
           // this.vendor_skuTableAdapter.Fill(this.dataSet3.vendor_sku);
            // TODO: This line of code loads data into the 'dataSet2.m_pe_maintain_lead_time' table. You can move, or remove it, as needed.
           // this.m_pe_maintain_lead_timeTableAdapter.Fill(this.dataSet2.m_pe_maintain_lead_time);

        }

        private void ultraGrid1_AfterRowActivate(object sender, EventArgs e)
        {
          
        }

        private void ultraGrid1_AfterSelectChange(object sender, Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs e)
        {
            txtFactLeadTime.Value = 0;
            if (this.ultraGrid1.Selected.Rows.Count > 0)
            {
                comboATFOB.Text = ultraGrid1.Selected.Rows[0].Cells[6].Value.ToString();
                comboATShipMeth.Text = ultraGrid1.Selected.Rows[0].Cells[7].Value.ToString();
                this.txtATLeadTime.Text = ultraGrid1.Selected.Rows[0].Cells[8].Value.ToString();
                comboNLFOB.Text = ultraGrid1.Selected.Rows[0].Cells[9].Value.ToString();
                comboNLShipMeth.Text = ultraGrid1.Selected.Rows[0].Cells[10].Value.ToString();
                this.txtNLLeadTime.Text = ultraGrid1.Selected.Rows[0].Cells[11].Value.ToString();
                comboSGFOB.Text = ultraGrid1.Selected.Rows[0].Cells[12].Value.ToString();
                comboSGShipMeth.Text = ultraGrid1.Selected.Rows[0].Cells[13].Value.ToString();
                this.txtSGLeadTime.Text = ultraGrid1.Selected.Rows[0].Cells[14].Value.ToString();
                if (ultraGrid1.Selected.Rows[0].Cells[5].Value.ToString()!="")
                    txtFactLeadTime.Value = Convert.ToInt32(ultraGrid1.Selected.Rows[0].Cells[5].Value.ToString());
                if (ultraGrid1.Selected.Rows[0].Cells[18].Value.ToString() == "Y")
                {
                    comboATFOB.Enabled = false;
                    comboATShipMeth.Enabled = false;
                    txtATLeadTime.Enabled = false;
                    comboNLFOB.Enabled = false;
                    comboNLShipMeth.Enabled = false;
                    txtNLLeadTime.Enabled = false;
                    comboSGFOB.Enabled = false;
                    comboSGShipMeth.Enabled = false;
                    lblVendorQuotes.Visible = true;
                    txtFactLeadTime.Enabled = false;
                    this.txtBoxesperpallet.Enabled = false;
                    this.txtQtyperbox.Enabled = false;
                    this.txtOrdMult.Enabled = false;
                    txtSGLeadTime.Enabled = false;
                }
                else
                {                   
                        comboATFOB.Enabled = true;
                        comboATShipMeth.Enabled = true;
                        txtATLeadTime.Enabled = true;
                        comboNLFOB.Enabled = true;
                        comboNLShipMeth.Enabled = true;
                        txtNLLeadTime.Enabled = true;
                        comboSGFOB.Enabled = true;
                        comboSGShipMeth.Enabled = true;
                        lblVendorQuotes.Visible = false;
                        txtFactLeadTime.Enabled = true;
                        txtSGLeadTime.Enabled = true;
                        this.txtBoxesperpallet.Enabled = true;
                        this.txtQtyperbox.Enabled = true;
                        this.txtOrdMult.Enabled = true;
                }

            }
        }
      //  public string ATFOB;
        public string ATShipMeth;
        private void comboATFOB_RowSelected(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
        {
            ATFOB = comboATFOB.SelectedText;
            loadATEditShipMethod(dsCombo.Tables[0]);
           // getLeadTime();          
        }
       
        private void ultraButton3_Click(object sender, EventArgs e)
        {
            this.ultraGrid1.PerformAction(UltraGridAction.Copy);
        }

        private void ultraButton4_Click(object sender, EventArgs e)
        {
            ultraGrid1.DisplayLayout.Bands[0].AddNew();
            ultraGrid1.Rows[ultraGrid1.Rows.Count - 1].Selected = true;
            this.ultraGrid1.PerformAction(UltraGridAction.Paste);
        }

        private void ultraButton5_Click(object sender, EventArgs e)
        {
            ultraGrid1.DisplayLayout.Bands[0].AddNew();
           // ultraGrid1.DisplayLayout.Override.AllowAddNew = AllowAddNew.FixedAddRowOnBottom;
           // ultraGrid1.DisplayLayout.Override.AllowAddNew = AllowAddNew.FixedAddRowOnTop;
           // ultraGrid1.DisplayLayout.Override.AllowAddNew = AllowAddNew.TemplateOnBottom;
           // ultraGrid1.DisplayLayout.Override.AllowAddNew = AllowAddNew.TemplateOnTop;
        }

        private void ultraGrid1_AfterRowInsert(object sender, RowEventArgs e)
        {
            ultraGrid1.DisplayLayout.Override.AllowAddNew = AllowAddNew.Yes;
            this.ultraGrid1.PerformAction(UltraGridAction.Paste);
        }
        public string ATComboShipMeth="";
        public string NLComboShipMeth = "";
        private string ATFOB="";
        private string NLFOB="";
        private void comboATShipMeth_RowSelected(object sender, RowSelectedEventArgs e)
        {
            ATComboShipMeth = comboATShipMeth.SelectedText;
            getATLeadTime();
        }

        public void getATLeadTime()
        {
            if (ATComboShipMeth != "" && this.comboATFOB.SelectedText.ToString() != "")
            {
                DataRow[] dr = dsCombo.Tables[0].Select(" site='AT' and ship_meth='" + ATComboShipMeth + "' and FOB='" + this.comboATFOB.SelectedText.ToString() + "'");
                if (dr.Length>0)
                       txtATLeadTime.Text = dr[0]["lead_time"].ToString();
            }
        }
        public void getNLLeadTime()
        {
            if (comboNLShipMeth.SelectedText.ToString() != "" && this.comboNLFOB.SelectedText.ToString() != "")
            {
                DataRow[] dr = dsCombo.Tables[0].Select(" site='NL' and ship_meth='" + comboNLShipMeth.SelectedText.ToString() + "' and FOB='" + this.comboNLFOB.SelectedText.ToString() + "'");
                if (dr.Length > 0)
                    txtNLLeadTime.Text = dr[0]["lead_time"].ToString();
            }
        }
       
        private void comboNLFOB_RowSelected(object sender, RowSelectedEventArgs e)
        {
            NLFOB = comboNLFOB.SelectedText;
            loadNLEditShipMethod(dsCombo.Tables[0]);
          //getLeadTime();   
        }

        private void comboNLShipMeth_RowSelected(object sender, RowSelectedEventArgs e)
        {
            getNLLeadTime(); 
        }

        private void loadNLEditShipMethod(DataTable dt)
        {
            if (this.comboNLFOB.SelectedText.ToString()!="")
            {
                DataTable dtShipMeth = dt.Select("site='NL' and FOB='" + this.comboNLFOB.SelectedText.ToString() + "'").CopyToDataTable().DefaultView.ToTable(true, "ship_meth");
                DataView dv = dtShipMeth.DefaultView;
                dv.Sort = "ship_meth";
                dtShipMeth = dv.ToTable();
                this.comboNLShipMeth.DataSource = dtShipMeth;
                comboNLShipMeth.DisplayMember = "ship_meth";
                comboNLShipMeth.ValueMember = "ship_meth";
            }
        }

        private void loadSGEditShipMethod(DataTable dt)
        {
            if (this.comboSGFOB.SelectedText.ToString() != "")
            {
                DataTable dtShipMeth = dt.Select("site='SG' and FOB='" + this.comboSGFOB.SelectedText.ToString() + "'").CopyToDataTable().DefaultView.ToTable(true, "ship_meth");
                DataView dv = dtShipMeth.DefaultView;
                dv.Sort = "ship_meth";
                dtShipMeth = dv.ToTable();
                this.comboSGShipMeth.DataSource = dtShipMeth;
                comboSGShipMeth.DisplayMember = "ship_meth";
                comboSGShipMeth.ValueMember = "ship_meth";
            }
        }

        public void getSGLeadTime()
        {
            if (comboSGShipMeth.SelectedText.ToString() != "" && this.comboSGFOB.SelectedText.ToString() != "")
            {
                DataRow[] dr = dsCombo.Tables[0].Select(" site='SG' and ship_meth='" + comboSGShipMeth.SelectedText.ToString() + "' and FOB='" + this.comboSGFOB.SelectedText.ToString() + "'");
                if (dr.Length > 0)
                    txtSGLeadTime.Text = dr[0]["lead_time"].ToString();
            }
        }

        private void comboSGShipMeth_RowSelected(object sender, RowSelectedEventArgs e)
        {
            getSGLeadTime(); 
        }

        private void comboSGFOB_RowSelected(object sender, RowSelectedEventArgs e)
        {  
            loadSGEditShipMethod(dsCombo.Tables[0]);
        }

        private void cmbATFOB_RowSelected(object sender, RowSelectedEventArgs e)
        {
            ATFOB = cmbATFOB.SelectedText;
            if (cmbATFOB.SelectedText.ToString() != "")
            {
                DataTable dtShipMeth = dsCombo.Tables[0].Select("site='AT' and [FOB]='" + cmbATFOB.SelectedText.ToString() + "'").CopyToDataTable().DefaultView.ToTable(true, "ship_meth");
                DataView dv = dtShipMeth.DefaultView;
                dv.Sort = "ship_meth";
                dtShipMeth = dv.ToTable();
                this.cmbATShipMeth.DataSource = dtShipMeth;
                cmbATShipMeth.DisplayMember = "ship_meth";
                cmbATShipMeth.ValueMember = "ship_meth";
            }
        }

        private void cmbNLFOB_RowSelected(object sender, RowSelectedEventArgs e)
        {
            if (cmbNLFOB.SelectedText.ToString() != "")
            {
                DataTable dtShipMeth = dsCombo.Tables[0].Select("site='NL' and [FOB]='" + cmbNLFOB.SelectedText.ToString() + "'").CopyToDataTable().DefaultView.ToTable(true, "ship_meth");
                DataView dv = dtShipMeth.DefaultView;
                dv.Sort = "ship_meth";
                dtShipMeth = dv.ToTable();
                this.cmbNLShipMeth.DataSource = dtShipMeth;
                cmbNLShipMeth.DisplayMember = "ship_meth";
                cmbNLShipMeth.ValueMember = "ship_meth";
            }
        }

        private void cmbSGFOB_RowSelected(object sender, RowSelectedEventArgs e)
        {
            if (cmbSGFOB.SelectedText.ToString() != "")
            {
                DataTable dtShipMeth = dsCombo.Tables[0].Select("site='SG' and [FOB]='" + cmbSGFOB.SelectedText.ToString() + "'").CopyToDataTable().DefaultView.ToTable(true, "ship_meth");
                DataView dv = dtShipMeth.DefaultView;
                dv.Sort = "ship_meth";
                dtShipMeth = dv.ToTable();
                this.cmbSGShipMeth.DataSource = dtShipMeth;
                cmbSGShipMeth.DisplayMember = "ship_meth";
                cmbSGShipMeth.ValueMember = "ship_meth";
            }
        }

        private void ultraButton1_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                string part_num = ultraGrid1.Selected.Rows[0].Cells[3].Value.ToString();
                string sqlStr = " update epicor.mimdist.dbo.part_price set m_factory_lt='" + txtFactLeadTime.Value.ToString() + "',  fob_at='" + comboATFOB.Text.ToString() + "',fob_nl='" + comboNLFOB.Text.ToString() + "', fob_sg='" + comboSGFOB.Text.ToString() + "', ship_meth_at='" + this.comboATShipMeth.Text.ToString() + "', ship_meth_nl='" + comboNLShipMeth.Text.ToString() + "', ship_meth_sg='" + comboSGShipMeth.Text.ToString() + "', m_transit_lt_at='" + txtATLeadTime.Text.ToString() + "', m_transit_lt_nl='" + txtNLLeadTime.Text.ToString() + "', m_transit_lt_sg='" + txtSGLeadTime.Text.ToString() + "', ord_mult='" + this.txtOrdMult.Text.ToString() + "', boxes_per_pallet='" + this.txtBoxesperpallet.Text.ToString() + "', qty_per_box='" + this.txtQtyperbox.Text.ToString() + "'     where part_no='" + part_num + "'  select @@error";
                SqlConnection cn = new SqlConnection(constrEpicor);
                cn.Open();
                SqlCommand cmd = new SqlCommand(sqlStr, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                cn.Close();

                if (ds.Tables[0].Rows[0][0].ToString() == "0")
                    MessageBox.Show("Updated successfully");
                loadGrid();
                Cursor.Current = Cursors.Default;
            }
            catch { }


        }

        private void ultraButton2_Click(object sender, EventArgs e)
        {
            clearAll();
        }

        public void clearAll()
        {
            comboATFOB.Value=null;
      }

        private void ultraButton6_Click(object sender, EventArgs e)
        {
           // MessageBox.Show(ultraNumericEditor1.Value.ToString());
            cmbManufacturer.Text = "";
           this.cmbATFOB.Text = "";
           this.cmbNLFOB.Text = "";
           this.cmbSGFOB.Text = "";
           this.cmbNLShipMeth.Text="";
           this.cmbATShipMeth.Text = "";
           this.cmbSGShipMeth.Text = "";
           this.cmbPartType.Text = "";
           

        }

        private void comboATShipMeth_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {

        }

        private void comboATFOB_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {

        }
     }

    }

