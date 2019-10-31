using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

// ARCHIVED - DO NOT USE
namespace ExcelLib
{
    public class Workbook
    {
        #region Members
        private string _dir, _filePath, _fileName, _fileExtension;
        private readonly Excel.Sheets _sheets;
        private readonly Excel.Application _exl;
        private readonly Excel._Workbook _book;
        #endregion

        #region Constructor and Finalizer
        public Workbook(string fullPath)
        {
            // Make sure the file exists before moving forward
            if (!File.Exists(fullPath)) { throw new FileNotFoundException(); }
            
            // Parse the path out into it's parts
            Directory       = Path.GetDirectoryName(fullPath);
            FileName        = Path.GetFileNameWithoutExtension(fullPath);
            FileExtension   = Path.GetExtension(fullPath);
            FilePath        = Path.GetFullPath(fullPath);
            Console.WriteLine("here");

            // Initialize Excel Members
            _exl = new Excel.Application();             // Initialize Excel Application
            exlApp.Visible = true;                      // Set Excel app visible
            //_book = exlApp.Workbooks.Open(FilePath,0, false, 5,"","", false, Excel.XlPlatform.xlWindows,"",true,false,0,true,false,false);
            _book = exlApp.Workbooks.Open(FilePath);
            _sheets = Book.Worksheets;
            //Spreadsheet s = new Spreadsheet((Excel.Worksheet)Sheets.get_Item("Sheet1"));
            Excel.Worksheet sheet = (Excel.Worksheet)Sheets.get_Item("Sheet1");
            sheet.Copy();
            s.Activate();
            //Excel.Range exlRange = s.get_Range("A2", "B20");
        }
        #endregion

        #region Functions

        #region Getters and Setters
        public string Directory
        {
            get { return this._dir; }
            set { this._dir = value; }
        }
        public string FilePath
        {
            get { return this._filePath; }
            set { this._filePath = value; }
        }
        public string FileName
        {
            get { return this._fileName; }
            set { this._fileName = value; }
        }
        public string FileExtension
        {
            get { return this._fileExtension; }
            set { this._fileExtension = value; }
        }
        public Excel.Sheets Sheets
        {
            get { return this._sheets; }
        }
        public Excel.Application exlApp
        {
            get { return _exl; }
        }
        public Excel._Workbook Book
        {
            get { return _book; }
        }
        #endregion
        #endregion
    }
}
