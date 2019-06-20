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
        IList<string> _companiesFromSetting;
        Dictionary<string, CompanyModel> _companyModels;
        IList<ErrorModel> _errorModels = new List<ErrorModel>();

        public Form1()
        {
            InitializeComponent();
            //InitializeDataStructures();
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
                txtLog.AppendText("Please make sure you have Companies.txt and predefine some companies.");
            }
            else
            {
                string text = File.ReadAllText(path);
                string[] companies = text.Split(',');
                foreach (var company in companies)
                {
                    _companiesFromSetting.Add(company.Trim());
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
            //build up company list
            _companiesFromSetting = new List<string>();
            _companyModels = new Dictionary<string, CompanyModel>();
            ReadCompaniesFromFile();
            //coner case check
            if (_companiesFromSetting.Count == 0)
                return;
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
            IList<FileInfo> files = new List<FileInfo>();
            GetFilesByBreadthFirstSearch(this.txtDir.Text, files);

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
            txtLog.AppendText($"\r\nJob done. Files searched: {files.Count}");

            //enable actions
            btnSubmit.Enabled = true;
            btnReset.Enabled = true;
        }

        private void GetFilesByBreadthFirstSearch(string dir, IList<FileInfo> files)
        {
            Queue<DirectoryInfo> queue = new Queue<DirectoryInfo>();
            queue.Enqueue(new DirectoryInfo(dir));
            while (queue.Count > 0)
            {
                DirectoryInfo curDir = queue.Dequeue();
                IList<FileInfo> curFiles = curDir.GetFiles("*.doc");
                foreach (FileInfo file in curFiles)
                {
                    files.Add(file);
                }
                DirectoryInfo[] subDirs = curDir.GetDirectories();
                foreach (DirectoryInfo subDir in subDirs)
                {
                    queue.Enqueue(subDir);
                }
            }
        }

        private string GetContractCompanyNameByFileName(string fileName)
        {
            foreach(string company in _companiesFromSetting)
            {
                if (fileName.Contains(company))
                    return company;
            }
            return DEFAULT_COMPANY + _defaultCompanyCount++;
        }

        private async System.Threading.Tasks.Task ReadDoc(string filePath, string fileName, string key, string company)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            try
            {
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
                    }
                }
                _companyModels[company].FileModels.Add(fileModel);
                docs.Close();
            }
            catch (Exception ex)
            {
                _errorModels.Add(new ErrorModel(filePath, ex.Message));
            }
            finally
            {
                word.Quit();
            }
        }

        private void WriteToExcel(string filePath)
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            _Workbook xlWorkbook = xlApp.Workbooks.Add(Type.Missing);
            //report
            _Worksheet xlReportSheet = xlWorkbook.Sheets[1];
            xlReportSheet.Name = "Report";
            int row = 1;
            xlReportSheet.Cells[row++, 1] = $"Error Count: {_errorModels.Count}";
            xlReportSheet.Cells[row, 1] = "File Name";
            xlReportSheet.Cells[row, 2] = "Error Message";
            row++;
            foreach (ErrorModel error in _errorModels)
            {
                xlReportSheet.Cells[row, 1] = error.FileName;
                xlReportSheet.Cells[row, 2] = error.ErrorMessage;
                row++;
            }

            //data
            var xlSheets = xlWorkbook.Sheets as Sheets;
            _Worksheet xlWorksheet = (Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
            xlWorksheet.Name = "Data";
            xlWorksheet.Cells[_lastRow, 1] = "Company";
            xlWorksheet.Cells[_lastRow, 3] = "File Name";
            xlWorksheet.Cells[_lastRow, 2] = "Amendment Files";
            xlWorksheet.Cells[_lastRow, 4] = "Paragraph";
            _lastRow++;
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
                _lastRow++;
            }

            xlWorkbook.SaveAs(filePath, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlWorkbook.Close();
            xlApp.Quit();
        }

    }
}
