using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;

namespace Version3
{
    public partial class ManualAdjustment : Form
    {
        public ManualAdjustment()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
          string sqlstr="";
                           /* sqlstr = sqlstr + "exec [sp_updateBTSForMT] '" + var_auto_id.ToString() + "','" + var_status + "','" + var_tran_id + "','" 
            '+ var_from_loc + "','" + var_to_loc + "','" + var_from_proj + "','" + var_to_proj + "','" + var_tran_date + "','" + var_part_no + "','"
                            * + var_tran_type + "','" + var_tran_source + "','" + var_tran_source_num + "','" + var_tran_source_line + "','" + 
                            * var_serial_num + "','" + var_qty + "','" 
            '+ var_reference + "','" + var_location + "','0','0','" + var_EDIreference + "','" + var_EDIStatus + "'";*/
          if (cmbFromLocation1.Text == "" || txtGPN1.Text.Trim() == "" || txtGPN2.Text.Trim() == "" || txtQty1.Value.ToString() == "0" || txtQty1.Value.ToString() == "" || cmbPlant1.Text == "ALL")
          {
              MessageBox.Show("Validation Error : Please enter all the mandatory data !!");
          }
          else
          {


              sqlstr = sqlstr + "exec [sp_updateBTSForADJ_GPNChange] '"
                  + txtTranID1.Value + "','"
                 
                  + cmbFromLocation1.Text + "','"
                   + lblToLocation1.Text + "','"
                  + txtFromProjCode1.Text + "','"
                  + txtToProjCode1.Text + "','"
                  + System.DateTime.Now + "','"
                  + txtGPN1.Text + "','MATERIAL TRANSFER','MT','"
                  + txtTranSourceNum1.Text
                  + "','1','','"
                  + txtQty1.Value + "','','"
                  + cmbPlant1.Text + "','"

                 + txtTranID2.Value + "','"
                
                 + lblFromLoc2.Text + "','"
                  + cmbToLocation2.Text + "','"
                 + txtFromProjCode2.Text + "','"
                 + txtToProjCode1.Text + "','"
                 + System.DateTime.Now + "','"
                 + txtGPN2.Text + "','MATERIAL TRANSFER','MT','"
                 + txtTranSourceNum2.Text + "','1','','"
                 + txtQty2.Value + "','','"
                 + cmbPlant2.Text + "'";
              UpdateRow(sqlstr);
              this.Close();
          }
        }

          public void UpdateRow(string sql)
        {           
                    string constr = ConfigurationManager.ConnectionStrings["MVIntConnectionString"].ConnectionString;
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
                clerALL();
                BTSForADJ bts = new BTSForADJ();
                bts.loadgrid();
                this.Close();
            }

            catch (Exception exp)
            {
                string error = exp.Message.ToString();
                // tran.Rollback();
                MessageBox.Show("Failed to Update. Please check the data");
            }
          //  loadgrid();
            cn.Close();
        }

        private void txtToLoc1_Leave(object sender, EventArgs e)
        {
            //txtFromLoc2.Text=txtToLoc1.Text;
        }

        private void txtQty1_Leave(object sender, EventArgs e)
        {
        txtQty2.Value=txtQty1.Value;
        }
        private void loadDropDown()
        {
          DataTable dtSiteFilter = new DataTable();

            dtSiteFilter.Columns.Add("Plant");
            dtSiteFilter.Rows.Add("ALL");
            dtSiteFilter.Rows.Add("FB");
            dtSiteFilter.Rows.Add("NF");



            cmbPlant1.DataSource = dtSiteFilter;
            cmbPlant1.DisplayMember = "Plant";
            cmbPlant1.ValueMember = "Plant";
            cmbPlant1.SelectedText = "ALL";

              cmbPlant2.DataSource = dtSiteFilter;
              cmbPlant2.DisplayMember = "Plant";
              cmbPlant2.ValueMember = "Plant";
            cmbPlant2.SelectedText = "ALL";



              DataTable dtFromToLoc = new DataTable();
            dtFromToLoc.Columns.Add("Location");
            dtFromToLoc.Rows.Add("CMMIMN2P");
            dtFromToLoc.Rows.Add("CMMIMN2PK");
    
            this.cmbFromLocation1.DataSource = dtFromToLoc;
        }

        private void cmbPlant1_Leave(object sender, EventArgs e)
        {
            cmbPlant2.Text=cmbPlant1.Text;
        }

        private void ultraCombo1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
        
        }

        private void cmbToLocation1_ListChanged(object sender, Infragistics.Win.ValueListChangedEventArgs e)
        {
            
        }

        private void ManualAdjustment_Load(object sender, EventArgs e)
        {
            this.loadDropDown();
            
        }

        private void cmbToLocation1_ValueChanged(object sender, EventArgs e)
        {
           // cmbToLocation2.Text = cmbFromLocation1.Text;
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void clerALL()
        {
            txtTranID1.Value = 0;
            txtTranID2.Value= 0;
            txtGPN1.Text = "";
            txtGPN2.Text = "";
            txtQty1.Value = 0;
            txtQty2.Value = 0;
            txtToProjCode1.Text = "MAIN";
            txtFromProjCode1.Text = "MAIN";
            txtToProjCode2.Text = "MAIN";
            txtFromProjCode2.Text = "MAIN";
            cmbFromLocation1.Text = "";
            cmbToLocation2.Text = "";
            cmbPlant1.Text = "";
            txtTranSourceNum1.Text = "";
            txtTranSourceNum2.Text = "";
            cmbPlant2.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            clerALL();
        }

        private void cmbFromLocation1_ValueChanged(object sender, EventArgs e)
        {
            cmbToLocation2.Text = cmbFromLocation1.Text;
        }




        }
    }

