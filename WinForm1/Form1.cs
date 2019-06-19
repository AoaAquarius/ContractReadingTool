using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using WinForm1.Models;

namespace WinForm1
{
    public partial class Form1 : Form
    {

        int _lastRow = 1;
        string _excelPath;
        readonly string DEFAULT_COMPANY = "Unknown";
        int _defaultCompanyCount = 1;
        IList<string> _companies = new List<string>();
        Dictionary<string, CompanyModel> _companyModels = new Dictionary<string, CompanyModel>();

        public Form1()
        {
            InitializeComponent();
            InitializeDataStructures();
        }

        private void InitializeDataStructures()
        {
            ReadCompaniesFromFile();
        }

        private void ReadCompaniesFromFile()
        {
            string fileName = "Companies.txt";
            string path = $"{Directory.GetCurrentDirectory()}\\{fileName}";
            if (!File.Exists(path))
            {
                MessageBox.Show("Please make sure you have Companies.txt.");
                System.Windows.Forms.Application.Exit();
            }
            else
            {
                string text = File.ReadAllText(path);
                string[] companies = text.Split(',');
                foreach (var company in companies)
                {
                    _companies.Add(company.Trim());
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            this.txtDir.Text = "";
            //this.txtDir.Text = "C:\\Users\\sdeng\\Desktop\\DataApiTest";
            this.txtKey.Text = "";
            //this.txtKey.Text = "Database";
            this.txtLog.Text = "";
        }

        private async void btnSubmit_Click(object sender, EventArgs e)
        {
            //coner case check
            if (String.IsNullOrWhiteSpace(this.txtDir.Text))
            {
                txtLog.AppendText($"\r\nPlease enter a valid directory.");
                return;
            }
            if (String.IsNullOrWhiteSpace(this.txtKey.Text))
            {
                txtKey.AppendText($"\r\nPlease enter a valid key.");
                return;
            }
            
            //disable actions
            btnSubmit.Enabled = false;
            btnReset.Enabled = false;

            //calculate output file path
            _excelPath = txtDir.Text + "\\output.xlsx";

            //preset all data
            DirectoryInfo d = new DirectoryInfo(this.txtDir.Text);
            FileInfo[] files = d.GetFiles("*.doc");
            InitialExcel(_excelPath);

            //deal with files
            foreach (FileInfo file in files)
            {
                txtLog.AppendText($"\r\nReading {file.Name}...\r\n");
                string company = GetContractCompanyNameByFileName(file.Name);
                if (!_companyModels.ContainsKey(company))
                {
                    _companyModels.Add(company, new CompanyModel());
                }
                if (file.Name.ToLower().Contains("sow"))
                {
                    _companyModels[company].AmendmentFileNames.Add(file.FullName);
                }
                else
                {
                    await ReadDoc(file.FullName, file.Name, this.txtKey.Text, company);
                }
                txtLog.AppendText($"\r\nFinish {file.Name}...");
                txtLog.AppendText("\r\n****************************************************************************************\r\n");
            }
            WriteToExcel(_excelPath);
            txtLog.AppendText("\r\nJob done.");

            //enable actions
            btnSubmit.Enabled = true;
            btnReset.Enabled = true;
        }

        private string GetContractCompanyNameByFileName(string fileName)
        {
            foreach(string company in _companies)
            {
                if (fileName.Contains(company))
                    return company;
            }
            return DEFAULT_COMPANY + _defaultCompanyCount++;
        }

        private async System.Threading.Tasks.Task ReadDoc(string filePath, string fileName, string key, string company)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object miss = System.Reflection.Missing.Value;
            object path = filePath;
            object readOnly = true;
            Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            int paraCount = 0;
            FileModel fileModel = new FileModel();
            fileModel.FileName = fileName;
            for (int i = 0; i < docs.Paragraphs.Count; i++)
            {
                string cur = docs.Paragraphs[i + 1].Range.Text.ToString();
                if (!ckbCaseSensitive.Checked)
                {
                    cur = cur.ToLower();
                    key = key.ToLower();
                }
                if (cur.Contains(key))
                {
                    paraCount++;
                    txtLog.AppendText($" \r\n{paraCount}. " + docs.Paragraphs[i + 1].Range.Text.ToString() + "\r\n");
                    fileModel.Paragraphs.Add(docs.Paragraphs[i + 1].Range.Text.ToString());
                    //WriteToExcel(_excelPath, fileName, docs.Paragraphs[i + 1].Range.Text.ToString(), _fileCount, paraCount);
                }
            }
            _companyModels[company].FileModels.Add(fileModel);
            docs.Close();
            word.Quit();
        }

        private void InitialExcel(string filePath)
        {
            //Start Excel and get Application object.
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();

            //Get a new workbook.
            _Workbook xlWorkbook = oXL.Workbooks.Add("");
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            xlWorksheet.Cells[_lastRow, 1] = "Company";
            xlWorksheet.Cells[_lastRow, 3] = "File Name";
            xlWorksheet.Cells[_lastRow, 2] = "Amendment Files";
            xlWorksheet.Cells[_lastRow, 4] = "Paragraph";
            _lastRow++;

            xlWorkbook.SaveAs(filePath, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            xlWorkbook.Close();
        }

        //private void WriteToExcel(string filePath, string fileName, string paragraph, int fileCount, int paraCount)
        //{
        //    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //    _Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
        //    _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        //    xlWorksheet.Cells[_lastRow, 1] = $"{fileCount}.{paraCount}";
        //    xlWorksheet.Cells[_lastRow, 2] = fileName;
        //    xlWorksheet.Cells[_lastRow, 3] = paragraph;
        //    _lastRow++;
        //    xlWorkbook.Save();
        //    xlWorkbook.Close();
        //}

        private void WriteToExcel(string filePath)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            _Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            foreach (KeyValuePair<string, CompanyModel> entry in _companyModels)
            {
                xlWorksheet.Cells[_lastRow, 1] = entry.Key;

                //append files
                int paraRow = _lastRow;
                foreach (FileModel file in entry.Value.FileModels)
                {
                    xlWorksheet.Cells[paraRow, 3] = file.FileName;
                    foreach(string para in file.Paragraphs)
                        xlWorksheet.Cells[paraRow++, 4] = para;
                }

                //append amendments
                int amendRow = _lastRow;
                foreach(string amendment in entry.Value.AmendmentFileNames)
                    xlWorksheet.Cells[amendRow++, 2] = amendment;

                _lastRow = Math.Max(paraRow, amendRow);
            }
            xlWorkbook.Save();
            xlWorkbook.Close();
        }

    }
}
