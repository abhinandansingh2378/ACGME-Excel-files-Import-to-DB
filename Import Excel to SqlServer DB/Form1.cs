using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
                        string sheetName = fileName;

                        FileInfo file = new FileInfo(fileLocation);

                        if (!file.Exists)
                        {
                            throw new Exception("Error, file doesn't exists!");
                        }

                        ExcelRead(fileLocation);
                        //Fetching  Stunum that we took input from Excel  
                        fileLocations[i] = fileLocation;
                        fileNames[i] = fileName;
                        files[i] = file;
                        //importdatafromexcel(fileLocation);
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

    }
}
