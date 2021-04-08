using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Deployment.Application;
using System.Configuration;

using System.Security;
using System.Security.Cryptography;
using System.Runtime.Remoting.Contexts;
using System.Security.Principal;
using System.Web;
using System.Diagnostics;

//using Google.Apis.Tasks.v1;
//using Google.Apis.Tasks.v1.Data;

//using datarepeater;



namespace Version3
{
    public partial class PlanningMenu : Form
    {
        
        public PlanningMenu()
        {
            InitializeComponent();
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                ApplicationDeployment deploy = ApplicationDeployment.CurrentDeployment;
                //  status.Text = "TEST";
                //txtUserID.Text = ConfigurationSettings.AppSettings["userName"].ToString();
              //  this.StatusLabel.Text = deploy.CurrentVersion.ToString();
            }
            else
            {
               // this.StatusLabel.Text = "TESTING";
            }
        }

        private void lnkForecastMaint_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            Forecast fc = new Forecast();
            fc.Show();
            Cursor.Current = Cursors.Default;
        }

        private void lnlLblBreakDownBatch_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://mvreporting/SFRouteConfigMtnc/ShopFloorFlex/bin/ShopFloorFlex.html");
        }

        private void lnkBuildPlanMaint_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //System.Diagnostics.Process.Start("http://mvreporting.materialinmotion.com/BuildPlan/Flex/BuildPlanoldformat/bin-debug/BuildPlan.html");
        }

        private void lnkEfficiency_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           // System.Diagnostics.Process.Start("http://fbfile.materialinmotion.com/LocalAppDashboard/EfficiencyRating.html");
        }

      
        private void lnkGoogleSharedDocs_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://docs.google.com/");
        }

        private void lnkMenloPortal_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://web.menlolog.com/wms/");
        }

        private void RegexSearch_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            PartSearch ps = new PartSearch();
            ps.Show();
        }

        private void lnlShopFloorBatch_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            PartMaintenance pm = new PartMaintenance();
            pm.Show();
        }

        private void linkDemandTracker_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            RackDeliveryPlan dt = new RackDeliveryPlan();
            dt.Show();
            Cursor.Current = Cursors.Default;
        }
              
        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //Menu mnu = new Menu();
           // mnu.Location = new Point(200, 200);
           // mnu.Show();
        }

        private void Label7_Click(object sender, EventArgs e)
        {

        }

        private void LinkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           
            WindowsIdentity id = WindowsIdentity.GetCurrent();
            string domainUserName = id.Name.Replace(@"MATMOTIONSJ\", " ").Trim();
            System.Diagnostics.Process.Start("http://mimnet.materialinmotion.com/VarianceTracker/Index.aspx?username=" + domainUserName);
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            System.Diagnostics.Process.Start("http://mvreporting/GIGBuildPlanWebservice/GIGBuildPlan/bin/UpdateBuildID.html");
        }

    
        private void lnklblBuildProg_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
         //   System.Diagnostics.Process.Start("http://atfile.materialinmotion.com/localappdashboard/traybuilddashboard.html");
        }

        private void lnklblProdLine_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           //     System.Diagnostics.Process.Start("http://atfile.materialinmotion.com/LocalAppDashboard/BuildDashboard_ProductionLine.html");
        }

        private void lnklblMatTrans_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
          //  System.Diagnostics.Process.Start("http://atfile.materialinmotion.com/LocalAppDashboard/BuildDashboard_MaterialTransfer.html");
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
          //  System.Diagnostics.Process.Start("http://atfile.materialinmotion.com/LocalAppDashboard/RackbuildDashboard.html");
        }

     

        //private void lnkDekitForecastMaint_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        //{
        //    GlobalFunction gf = new GlobalFunction();
        //    string userid = "admin@materialinmotion.com";
        //    string password = "mimi1000";
        //    gf.createDocService(userid, password);
        //    gf.createSpreadsheetService(userid, password);
        //    docService = gf.docService;
        //    spreadsheetService = gf.spreadsheetService;
        //    DekitForecast fc = new DekitForecast(spreadsheetService, docService);
        //    fc.Show();    
        //}

        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        //    GlobalReport gb = new GlobalReport();
          //  gb.Show();
        }

        private void lnkShopFloorMRB_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
          
            Cursor.Current = Cursors.WaitCursor;
            VendorCommitApp vc = new VendorCommitApp();
            vc.Show();
            Cursor.Current = Cursors.Default;
          
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
           NPI_MRP npi = new NPI_MRP();
            npi.Show();
            Cursor.Current = Cursors.Default;
        }

        private void linkLabel2_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void PlanningMenu_Load(object sender, EventArgs e)
        {

        }

        private void linkLabel3_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //MRP_PROCESS_RUN mrp = new MRP_PROCESS_RUN();
          // mrp.Show();
        }

        private void Label39_Click(object sender, EventArgs e)
        {

        }

        private void label29_Click(object sender, EventArgs e)
        {

        }

        private void Label42_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://mvreporting/GIGBuildPlanWebservice/GIGBuildPlan/bin/menu.html");
        }

      

        private void linkLabel8_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {

            Cursor.Current = Cursors.WaitCursor;
         //   VendorCommit vc = new VendorCommit();
        //    vc.Show();
            Cursor.Current = Cursors.Default;

        }

        private void linkLabel9_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

          AssignedPart assPart = new AssignedPart();
            assPart.Show();

        }

        private void label15_Click(object sender, EventArgs e)
        {


        }

        //private void label33_Click(object sender, EventArgs e)
        //{
        //    GridTestFrom grd = new GridTestFrom();
        //    grd.Show();
        //}

        private void linkLabel2_LinkClicked_2(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void linkLabel10_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            VendorSKU frm = new VendorSKU();
            frm.Show();
           // ProcessStartInfo startInfo = new ProcessStartInfo();
           //  startInfo.FileName = "datarepeater.exe";
           //  startInfo.UseShellExecute = true;
           //  Process.Start(startInfo);
        
        }

        private void linkLabel3_LinkClicked_2(object sender, LinkLabelLinkClickedEventArgs e)
        {

            System.Diagnostics.Process.Start("https://vendor-portal.corp.google.com");

        }

        private void linkVendorPortal_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Vendor_Protal vp = new Vendor_Protal();
            vp.Show();
        }

        private void linkLabel12_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            RepeaterTest rep = new RepeaterTest();
            rep.Show();
        }

        private void linkLabel7_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            VendorTrackingNum tracking = new VendorTrackingNum();
           tracking.Show();
        }

        private void linkLabel13_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            VendorSplit split = new VendorSplit();
            split.Show();
        }

        private void linkLabel14_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BTSForPOReceipt frm = new BTSForPOReceipt();
            frm.Show();
        }

        private void linkLabel15_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BTS_Netgear frmNet = new  BTS_Netgear();
            frmNet.Show();
        }

        private void linkLabel16_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           Kanban kc = new Kanban();
            kc.Show();
        }

        private void linkLabel17_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           frmQuoteCopy quoteCopy = new frmQuoteCopy();
           quoteCopy.Show();
        }

        private void linkLabel18_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            
        }

        private void label44_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel11_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://atfile.materialinmotion.com/localappdashboard/rackbuilddashboard.html");
        }

        private void lnkUpdateInvoice_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmUpdateInvoice inv = new frmUpdateInvoice();
            inv.Show();
        }

        private void linkLabel19_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BuildStatusReport bld = new BuildStatusReport();
            bld.Show();
        }

        private void linkLabel20_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {          
          //  MOR_Request mor = new MOR_Request();
          //  mor.Show();
        }

        private void linkLabel21_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            GIGForecast gig = new GIGForecast();
            gig.Show();
        }

        private void linkLabel5_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Packaging_BOM bom = new Packaging_BOM();
            bom.Show();
        }

        private void linkLabel18_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
           
        }

        private void linkLabel18_LinkClicked_2(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Bacon_BOM bacon = new Bacon_BOM();
            bacon.Show();
        }

        private void linkLabel22_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            RepPartList part = new RepPartList();
            part.Show();
        }

        private void linkLabel23_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            RegionOption reg = new RegionOption();
            reg.Show();
        }

        private void linkLabel24_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FirstArticleForm firstArt = new FirstArticleForm();
           firstArt.Show();
        }

        private void linkLabel25_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           WB_BTSForm WB = new WB_BTSForm();
           WB.Show();
        }

        private void linkLabel27_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            KanbanPartAssign kb_part = new KanbanPartAssign();
            kb_part.Show();
        }

        private void linkLabel26_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
         //  KanbanOrderAssign kb_order = new KanbanOrderAssign();
          // kb_order.Show();
        }

        private void label57_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel28_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           NewBuildPlan nbp = new NewBuildPlan();
           nbp.Show();
        }

      /*  private void linkLabel29_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BTSMainForm bts = new BTSMainForm();
            bts.Show();
        }

        private void linkLabel30_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BTSNetgearForm btsNetgear = new BTSNetgearForm();
            btsNetgear.Show();
        }
        */
        private void linkLabel31_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("\\\\mvfile\\MV REPORTS\\Master Data & Forms\\PlanningExecutionForms\\VCAPP\\VCADemandSheet52Week.xltm");
          
        }

        private void linkLabel8_LinkClicked_2(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("\\\\mvfile\\MV REPORTS\\Master Data & Forms\\PlanningExecutionForms\\HomeMade MRP\\MRPExcel.xltm");
        }

        private void linkLabel32_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           MRP_Process mrp = new MRP_Process();
           mrp.Show();
        }

        private void linkLabel33_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           BTSForMT btsMT = new BTSForMT();
           btsMT.Show();
        }

        private void label62_Click(object sender, EventArgs e)
        {

        }

        private void label59_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel20_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BTSForMT btsMT = new BTSForMT();
            btsMT.Show();
        }

        private void linkLabel33_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BTSForMR btsMR = new BTSForMR();
            btsMR.Show();
        }

        private void linkLabel34_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BTSForAC btsAC = new BTSForAC();
            btsAC.Show();
        }

        private void linkLabel35_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BTSForDecom2 btsDecom = new BTSForDecom2();
            btsDecom.Show();
        }

        private void linkLabel36_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BTSForADJ btsAdj = new BTSForADJ();
            btsAdj.Show();

        }

        private void linkLabel37_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BTSErrorLog error = new BTSErrorLog();
            error.Show();
        }

        private void linkLabel38_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BTSNewBuildForAC btsBuild = new BTSNewBuildForAC();
            btsBuild.Show();
        }

        private void linkLabel39_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            IntegrationMTRequest mtreq = new IntegrationMTRequest();
            mtreq.Show();
        }

        private void linkLabel40_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ExecuteDisposition disp = new ExecuteDisposition();
            disp.Show();
        }

        private void Label45_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel15_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ImportInvoice inv = new ImportInvoice();
            inv.Show();
        }

        private void linkLabel16_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            VanillaCommit npicom = new VanillaCommit();
            npicom.Show();
        }

        private void lnklblVendorSplit_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

    }
}
 