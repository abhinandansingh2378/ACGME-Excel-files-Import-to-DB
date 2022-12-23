using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Import_Excel_to_SqlServer_DB
{

    
    public partial class Form1 : Form
    {

        string[] fileLocations;
        string[] fileNames;
        FileInfo[] files;
        string fileLocation = "";
        string fileName = "";
        DataTable dt = new DataTable();
        DataColumn dtColumn;
        DataRow myDataRow;
        string[] acgmeColumns = { "No.", "Program Name", "Program Name2",    "Address",  "Address2", "City", "Contact",  "Phone No.",    "Global Dimension 1 Code",  "Country/Region Code",  "Zip Code", "State",    "Email",    "Primary Contact No.",  "Vendor Sub Type",  "ACGME #",  "Residency",    "Non-Affiliated Hospital",  "State Code",   "Speciality",   "Extension",    "Primary Contact",  "Primary Contact Name", "Program Director", "Accreditation Status", "Effective Date",   "Clinical Rotation Exists"};
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //create openfileDialog Object
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            //open file format define PDF files
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Multiselect = true; //allow multiline selection at the file selection level
            openFileDialog1.InitialDirectory = @"D:"; //define the initial directory

            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK) //executing when file open
                {
                    fileLocations = new String[openFileDialog1.FileNames.Length];
                    fileNames = new String[openFileDialog1.FileNames.Length];
                    files = new FileInfo[openFileDialog1.FileNames.Length];


                    for (int i = 0; i < openFileDialog1.FileNames.Length; i++)
                    {

                        fileLocation = openFileDialog1.FileNames[i];

                        fileName = System.IO.Path.GetFileNameWithoutExtension(openFileDialog1.FileNames[i]).Replace(".pdf","");
                        string strConn = string.Empty;
                        fileName = "[" +fileName+"]";

                        FileInfo file = new FileInfo(fileLocation);

                        if (!file.Exists)
                        {
                            throw new Exception("Error, file doesn't exists!");
                        }

                        ExcelRead(fileLocation);
                        //Fetching  Stunum that we took input from Excel  
                        BulkInsertDataTable(fileName, dt);
                        textBox1.Text += fileName + " table is inserted to Acgme Data base " + "\n";
                        //importdatafromexcel(fileLocation);
                        dt.Reset();
                    }


                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error, Issue while selecting PDF file ! ");
            }
        }

        public void ExcelRead(string pathName)
        {
            try
            {
                dt.Reset();
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(pathName);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int columnCount = xlRange.Columns.Count;
                AcgmeColumns();
                for (int i = 2; i <= rowCount; i++)
                {
                    myDataRow = dt.NewRow();
                    for (int j = 1; j <= columnCount; j++)
                    {
                        
                        if (j == 1 )
                        {
                            myDataRow[acgmeColumns[j - 1]] = Convert.ToUInt64(xlRange.Cells[i, j].Value2);
                        }
                        else
                        {
                            myDataRow[acgmeColumns[j - 1]] = Convert.ToString(xlRange.Cells[i, j].Value2);
                        }
                        
                    }
                    dt.Rows.Add(myDataRow);
                }
                xlWorkbook.Close();
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Getting Error " + ex.ToString());
            }


        }

        public void AcgmeColumns()
        {
            dt.Columns.Add("No.");
            dt.Columns.Add("Program Name");
            dt.Columns.Add("Program Name2");
            dt.Columns.Add("Address");
            dt.Columns.Add("Address2");
            dt.Columns.Add("City");
            dt.Columns.Add("Contact");
            dt.Columns.Add("Phone No.");
            dt.Columns.Add("Global Dimension 1 Code");
            dt.Columns.Add("Country/Region Code");
            dt.Columns.Add("Zip Code");
            dt.Columns.Add("State");
            dt.Columns.Add("Email");
            dt.Columns.Add("Primary Contact No.");
            dt.Columns.Add("Vendor Sub Type");
            dt.Columns.Add("ACGME #");
            dt.Columns.Add("Residency");
            dt.Columns.Add("Non-Affiliated Hospital");
            dt.Columns.Add("State Code");
            dt.Columns.Add("Speciality");
            dt.Columns.Add("Extension");
            dt.Columns.Add("Primary Contact");
            dt.Columns.Add("Primary Contact Name");
            dt.Columns.Add("Program Director");
            dt.Columns.Add("Accreditation Status");
            dt.Columns.Add("Effective Date");
            dt.Columns.Add("Clinical Rotation Exists");
        }

        public static String CreateTableImport(string tablename)
        {
            return @"IF OBJECT_ID(N'dbo." + tablename + "',N'U') IS NOT NULL" + " DROP TABLE [dbo]." + tablename
            +" CREATE TABLE " + "[dbo]." +tablename
                   + "([ACGME No#] nvarchar(255),"
                   + "[Program Name] nvarchar(255),"
                   + "[Program Name2] nvarchar(255),"
                   + "[Address] nvarchar(255),"
                   + "[Address2] nvarchar(255),"
                   + "[City] nvarchar(255),"
                   + "[Contact] nvarchar(255),"
                   + "[Phone No#] nvarchar(255),"
                   + "[Global Dimension 1 Code] nvarchar(255),"
                   + "[Country/Region Code] nvarchar(255),"
                   + "[Zip Code] nvarchar(255),"
                   + "[State] nvarchar(255),"
                   + "[Email] nvarchar(255),"
                   + "[Primary Contact No#] nvarchar(255),"
                   + "[Vendor Sub Type] nvarchar(255),"
                   + "[ACGME #] nvarchar(255),"
                   + "[Residency] nvarchar(255),"
                   + "[Non-Affiliated Hospital] nvarchar(255),"
                   + "[State Code] nvarchar(255),"
                   + "[Speciality] nvarchar(255),"
                   + "[Extension] nvarchar(255),"
                   + "[Primary Contact] nvarchar(255),"
                   + "[Primary Contact Name] nvarchar(255),"
                   + "[Program Director] nvarchar(255),"
                   + "[Accreditation Status] nvarchar(255),"
                   + "[Effective Date] nvarchar(255),"
                   + "[Clinical Rotation Exists] nvarchar(255))";

        }

        public void BulkInsertDataTable(string tableName, DataTable dataTable)
        {
            try
            {
                string sqlConnectionString = "server=mea-dm;User ID = tricon; password = mea@1234; database = Acgme; connection reset = false";
                SqlConnection sqlconn = new SqlConnection(sqlConnectionString);
                sqlconn.Open();
                // create table if not exists 
                string createTableQuery = CreateTableImport(tableName);
                SqlCommand createCommand = new SqlCommand(createTableQuery, sqlconn);
                createCommand.ExecuteNonQuery();
                SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlconn, SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.FireTriggers |SqlBulkCopyOptions.UseInternalTransaction, null);
                bulkCopy.DestinationTableName = tableName;
                bulkCopy.WriteToServer(dataTable);
                sqlconn.Close();
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message.ToString());
            }
        }

    }
}
