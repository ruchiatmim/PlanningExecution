using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;

namespace Version3
{
    public partial class BTSErrorLog : Form
    {
        public BTSErrorLog()
        {
            InitializeComponent();
        }
        string filePath = "";
        private void button2_Click(object sender, EventArgs e)
        {
           DialogResult result= openFileDialog1.ShowDialog();
           if (result == DialogResult.OK) // Test result.
           {
                filePath = openFileDialog1.FileName;
                textBox1.Text = filePath;  
           }
          
        }


        static void CovertExcelToCsv(string excelFilePath, string csvOutputFile, int worksheetNumber = 1)
        {
           // if (!File.Exists(excelFilePath)) throw new FileNotFoundException(excelFilePath);
           // if (File.Exists(csvOutputFile)) throw new ArgumentException("File exists: " + csvOutputFile);

            // connection string
            var cnnStr = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\users\\rverma\\Documents\\PROD_MIM_plat_error_20160619020956 - Copy.xls;Extended Properties=Excel 12.0;");
            var cnn = new OleDbConnection(cnnStr);

            // get schema, then data
            var dt = new DataTable();
            try
            {
                cnn.Open();
                var schemaTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schemaTable.Rows.Count < worksheetNumber) throw new ArgumentException("The worksheet number provided cannot be found in the spreadsheet");
                string worksheet = schemaTable.Rows[worksheetNumber - 1]["table_name"].ToString().Replace("'", "");
                string sql = String.Format("select * from [{0}]", worksheet);
                var da = new OleDbDataAdapter(sql, cnn);
                da.Fill(dt);
            }
            catch (Exception e)
            {
                // ???
                //throw e;
            }
            finally
            {
                // free resources
                cnn.Close();
            }

            // write out CSV data
            using (var wtr = new StreamWriter(csvOutputFile))
            {
                foreach (DataRow row in dt.Rows)
                {
                    bool firstLine = true;
                    foreach (DataColumn col in dt.Columns)
                    {
                        if (firstLine) 
                        { 
                            wtr.Write(","); 
                         } 
                        
                        else { firstLine = false; }
                        var data = row[col.ColumnName].ToString().Replace("\"", "\"\"");
                        wtr.Write(String.Format("\"{0}\"", data));
                    }
                    wtr.WriteLine();
                }
            }
        }

             private void button1_Click(object sender, EventArgs e)
             {

                 filePath = textBox1.Text;//.Substring(textBox1.Text.LastIndexOf("/")+1,textBox1.Text.Length);
                 int strstart = textBox1.Text.ToString().Trim().LastIndexOf("\\") ;
                 string filename = filePath.Substring(strstart + 1, filePath.Length - (strstart + 1));
               //  string csvOutputFile = filename.Replace("xls", "csv");
                 //FileInfo file=new FileInfo(filePath);
                 System.IO.File.Copy(filePath, "\\\\mimsftp\\e$\\sftp\\msgGoogle\\MDS\\PROD\\inbound\\INTEGRATIONERRORREPORT\\ErrorReport\\" + filename, true);
             ////   file.CopyTo("\\\\mimsftp\\e$\\sftp\\msgGoogle\\PLATFORM\\PROD\\ERRORS\\" + csvOutputFile, true);
                // if (filename.Contains(".xls"))
                  //   ConvertExcelToCsv(filePath, "\\\\mimsftp\\e$\\sftp\\msgGoogle\\MDS\\PROD\\inbound\\INTEGRATIONERRORREPORT\\ErrorReport\\CSV File\\ErrorRepot.csv" + filename, 1);

                runJob();

              //   jobConnection = new SqlConnection("Data Source=(local);Initial Catalog=msdb;Integrated Security=SSPI");
             //    jobCommand = new SqlCommand("sp_start_job", jobConnection);
               //  jobCommand.CommandType = CommandType.StoredProcedure;  
  
             }
             string constr = ConfigurationManager.ConnectionStrings["MVReportingConnectionString"].ConnectionString;
             private void runJob()
             {
                 SqlConnection cn = new SqlConnection(constr);
                 cn.Open();
                 SqlTransaction tran = cn.BeginTransaction();
                 try
                 {
                     SqlCommand cmd = new SqlCommand(" EXEC msdb.dbo.sp_start_job 'SIErrorReportUpload' ", cn);
                     cmd.CommandTimeout = 0;
                     cmd.Transaction = tran;
                     cmd.ExecuteNonQuery();
                     tran.Commit();
                     MessageBox.Show("Job has been statrted", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 }
                 catch (Exception e)
                 {
                     MessageBox.Show("Exception to upload file");
                 }


             }

             private void ConvertExcelToCsv(string excelFilePath, string csvOutputFile, int worksheetNumber = 1)
             {
                 System.Data.OleDb.OleDbConnection cnn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';");
                         
                 // get schema, then data
                var dt = new DataTable();
                try {
                        cnn.Open();
                        var schemaTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        if (schemaTable.Rows.Count < worksheetNumber) throw new ArgumentException("The worksheet number provided cannot be found in the spreadsheet");
                        string worksheet = schemaTable.Rows[worksheetNumber-1 ]["table_name"].ToString().Replace("'", "");
                        string sql = String.Format("select * from [{0}]", worksheet);
                        var da = new OleDbDataAdapter(sql, cnn);
                        da.Fill(dt);
                    }
                catch (Exception e) 
                {
                     // ???
                    throw e;
                }
                finally 
                {
                    // free resources
                    cnn.Close();
                }

                using (var wtr = new StreamWriter(csvOutputFile))
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        bool firstLine = true;
                        foreach (DataColumn col in dt.Columns)
                        {
                            if (!firstLine) { wtr.Write(","); } else { firstLine = false; }
                            var data = row[col.ColumnName].ToString().Replace("\"", "\"\"");
                            wtr.Write(String.Format("\"{0}\"", data));
                        }
                        wtr.WriteLine();
                    }
                }


             }

             private void btnCancel_Click(object sender, EventArgs e)
             {
                 textBox1.Text = "";
             }
    }

      
}
   

