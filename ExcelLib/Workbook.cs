using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    public class Workbook
    {
        #region Fields

        #region Private
        private ExcelControl _parent;
        private Excel.Workbook _WBObj;
        #endregion

        #region Public
        public List<Spreadsheet> Worksheets { get; }
        public string Directory { get; set; }
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string FileExtension { get; set; }
        #endregion

        #endregion

        #region Constructor
        public Workbook(ExcelControl parent, Excel.Workbook WorkbookObj, string fullPath)
        {
            // Parse the path out into it's parts
            Directory = Path.GetDirectoryName(fullPath);
            FileName = Path.GetFileNameWithoutExtension(fullPath);
            FileExtension = Path.GetExtension(fullPath);
            FilePath = Path.GetFullPath(fullPath);
            Console.WriteLine("here");

            // Initialize fields
            _WBObj = WorkbookObj;
            _parent = parent;
            Worksheets = new List<Spreadsheet>();
            
            // Populate Spreadsheet List
            foreach(Excel.Worksheet sheet in _WBObj.Worksheets)
            {
                Worksheets.Add(new Spreadsheet(this, sheet));
            }

        }
        #endregion

        #region Accessors

        public ExcelControl getParentApp()
        {
            return _parent;
        }

        #endregion

        #region Methods
 
        #endregion
    }
}

