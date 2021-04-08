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
    public partial class Packaging_BOM : Form
    {
        string constr = ConfigurationManager.ConnectionStrings["epicorConnectionString"].ConnectionString;
        public Packaging_BOM()
        {
            InitializeComponent();
            loadData();
        }

        public void loadData()
        {

            DataSet ds = getBOM();
            ultraGridBOM.DataSource = ds.Tables[0];
            lblNumRecKan.Text = "Num of Records :"+  ds.Tables[0].Rows.Count.ToString();
            ultraGridBOM.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;

            Infragistics.Win.UltraWinGrid.UltraGridColumn c1 = ultraGridBOM.DisplayLayout.Bands[0].Columns["Plan Group"];
            c1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

            Infragistics.Win.UltraWinGrid.UltraGridColumn c2 = ultraGridBOM.DisplayLayout.Bands[0].Columns["Form Factor"];
            c2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

            Infragistics.Win.UltraWinGrid.UltraGridColumn c3 = ultraGridBOM.DisplayLayout.Bands[0].Columns["Platform"];
            c3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

            Infragistics.Win.UltraWinGrid.UltraGridColumn c4 = ultraGridBOM.DisplayLayout.Bands[0].Columns["Machine Type"];
            c4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

            Infragistics.Win.UltraWinGrid.UltraGridColumn c5 = ultraGridBOM.DisplayLayout.Bands[0].Columns["Publish Site"];
            c5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

            Infragistics.Win.UltraWinGrid.UltraGridColumn c6 = ultraGridBOM.DisplayLayout.Bands[0].Columns["Make Site"];
            c6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

            Infragistics.Win.UltraWinGrid.UltraGridColumn c7 = ultraGridBOM.DisplayLayout.Bands[0].Columns["Region"];
            c7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

            Infragistics.Win.UltraWinGrid.UltraGridColumn c8 = ultraGridBOM.DisplayLayout.Bands[0].Columns["PG BOM NUM"];
            c8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit; 


            DataView view = new DataView(ds.Tables[0]);
            DataTable distinctValuesPlanGroup = view.ToTable(true, "Plan Group");
            DataTable distinctValuesFormFactor = view.ToTable(true, "Form Factor");
            DataTable distinctValuesPlatform = view.ToTable(true, "Platform");
            DataTable distinctValuesMachineType = view.ToTable(true, "Machine Type");

            DataRow drplan = distinctValuesPlanGroup.NewRow();
            drplan[0] = "ALL";
          
            distinctValuesPlanGroup.Rows.InsertAt(drplan, 0);

            DataRow drFormFactor = distinctValuesFormFactor.NewRow();
            drFormFactor[0] = "ALL";
          
            distinctValuesFormFactor.Rows.InsertAt(drFormFactor, 0);

            DataRow drPlatform = distinctValuesPlatform.NewRow();
            drPlatform[0] = "ALL";
            
            distinctValuesPlatform.Rows.InsertAt(drPlatform, 0);

            DataRow drMachineType = distinctValuesMachineType.NewRow();
            drMachineType[0] = "ALL";
           
            distinctValuesMachineType.Rows.InsertAt(drMachineType, 0);
            
            comboMachineType.DataSource = distinctValuesMachineType;
            comboMachineType.DisplayMember = "Machine Type";
            comboMachineType.ValueMember = "Machine Type";

            
            comboPlatform.DataSource = distinctValuesPlatform;
            comboPlatform.DisplayMember = "Platform";
            comboPlatform.ValueMember = "Platform";
           
            ultraComboFormFactor.DataSource = distinctValuesFormFactor;
            ultraComboFormFactor.DisplayMember = "Form Factor";
            ultraComboFormFactor.ValueMember = "Form Factor";

            comboPlanGroup.DataSource = distinctValuesPlanGroup;
            comboPlanGroup.DisplayMember = "Plan Group";
            comboPlanGroup.ValueMember = "Plan Group";


        }

        public DataSet  getBOM()
        {
            string sqlQry = "SELECT [attrib_4] as [Plan Group]  ,[attrib_1]  as [Form Factor],[attrib_2] as Platform ,[attrib_3] as [Machine Type]  ,[Region] ,[publish_site] as [Publish Site] ,[make_site] as [Make Site]  ,[pg_bom_num] as [PG BOM NUM]  FROM [MIMDIST].[dbo].[d_offsets_region_to_asm] where view_filter='BP'";
    
            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlQry, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            return ds;
        }

        private void btnFilter_Click(object sender, EventArgs e)
        {
            filterData(); 
        }


        public void filterData()
        {
            string sqlQry = "SELECT [attrib_4] as [Plan Group]  ,[attrib_1]  as [Form Factor],[attrib_2] as Platform ,[attrib_3] as [Machine Type]  ,[Region] ,[publish_site] as [Publish Site] ,[make_site] as [Make Site]  ,[pg_bom_num]  as [PG BOM NUM]   FROM [MIMDIST].[dbo].[d_offsets_region_to_asm] where view_filter='BP'";


            if (comboPlatform.Text != "ALL" && comboPlatform.Text != "")
                sqlQry = sqlQry + " and attrib_2='" + comboPlatform.Text + "'";

            if (comboPlanGroup.Text != "ALL" && comboPlanGroup.Text != "")
                sqlQry = sqlQry + " and attrib_4='" + comboPlanGroup.Text + "'";

            if (ultraComboFormFactor.Text != "ALL" && ultraComboFormFactor.Text != "")
                sqlQry = sqlQry + " and attrib_1='" + ultraComboFormFactor.Text + "'";

            if (comboMachineType.Text != "ALL" && comboMachineType.Text != "")
                sqlQry = sqlQry + " and attrib_3='" + comboMachineType.Text + "'";


            SqlConnection cn = new SqlConnection(constr);
            cn.Open();
            SqlCommand cmd = new SqlCommand(sqlQry, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            cn.Close();
            ultraGridBOM.DataSource = ds.Tables[0];
            lblNumRecKan.Text = "Num of Records :" + ds.Tables[0].Rows.Count.ToString();
            ultraGridBOM.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;


        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            String sql = "";
            for (int i = 0; i < ultraGridBOM.Rows.Count; i++)
            {
                string attrib_1 = ultraGridBOM.Rows[i].Cells["Form Factor"].Value.ToString();
                string attrib_2 = ultraGridBOM.Rows[i].Cells["Platform"].Value.ToString();
                string attrib_3 = ultraGridBOM.Rows[i].Cells["Machine Type"].Value.ToString();
                string attrib_4 = ultraGridBOM.Rows[i].Cells["Plan Group"].Value.ToString();
                string region = ultraGridBOM.Rows[i].Cells["Region"].Value.ToString();
                string publish_site = ultraGridBOM.Rows[i].Cells["Publish Site"].Value.ToString();
                string make_site = ultraGridBOM.Rows[i].Cells["Make Site"].Value.ToString();
                string pg_bom_num = ultraGridBOM.Rows[i].Cells["PG Bom Num"].Value.ToString();

             sql = sql +  "  update   [MIMDIST].[dbo].[d_offsets_region_to_asm] set pg_bom_num='" + pg_bom_num + "' where   view_filter='BP' and make_site='" + make_site + "' and publish_site='" + publish_site + "' and region='" + region + "' and attrib_4='" + attrib_4 + "' and attrib_3='" + attrib_3 + "'  and attrib_2='" + attrib_2 + "' and attrib_1='" + attrib_1 + "'";
            }

            SqlConnection cn = new SqlConnection(constr);
             cn.Open();
             SqlTransaction tran = cn.BeginTransaction();
             try
             {                
                 SqlCommand cmd = new SqlCommand(sql, cn);
                 cmd.Transaction = tran;
                 cmd.ExecuteNonQuery();
                 tran.Commit();
                 MessageBox.Show("Data is saved successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 filterData();
             
             }
             catch (Exception exp)
             {
                 tran.Rollback();
                 MessageBox.Show("Not Saved. Please check the error" + exp.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

             }
         
             cn.Close();
             
        }

        private void Packaging_BOM_Load(object sender, EventArgs e)
        {

        }
    }
}
