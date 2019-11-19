using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        private Excel.Workbook _EditObj;
        private Excel.Window _WinObj;
        private bool isOpen;
        #endregion

        #region Public
        public List<Spreadsheet> Worksheets { get; }
        public string Directory { get; set; }
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string FileExtension { get; set; }
        public string TempPath { get; set; }
        public bool isShowing { get; private set; }
        #endregion

        #endregion

        #region Constructor
        public Workbook(ExcelControl parent, Excel.Workbook WorkbookObj, string fullPath = null)
        {
            // Check if path was given and if so, set all of the path fields 
            if (fullPath != null)
                setPath(fullPath);

            // Initialize fields
            _WBObj = WorkbookObj;
            _parent = parent;
            isShowing = false;
            Worksheets = new List<Spreadsheet>();
            
            // Populate Spreadsheet List
            foreach(Excel.Worksheet sheet in _WBObj.Worksheets)
            {
                Worksheets.Add(new Spreadsheet(this, sheet));
            }

        }

        public Workbook(ExcelControl parent, string fullPath)
        {
            // Initialize fields
            setPath(fullPath);
            _WBObj = parent.App.Workbooks.Open(fullPath, ReadOnly: true);
            _parent = parent;
            _WinObj = _WBObj.Windows[1];
            isShowing = false;
            isOpen = false;
            Worksheets = new List<Spreadsheet>();

            // Populate Spreadsheet List
            foreach (Excel.Worksheet sheet in _WBObj.Worksheets)
            {
                Worksheets.Add(new Spreadsheet(this, sheet));
            }
        }
        #endregion

        #region Accessors
        public ExcelControl Control
        {
            get
            {
                return _parent;
            }
        }
        public string Name
        {
            get
            {
                return FileName;
            }
        }
        public Spreadsheet ActiveSheet
        {
            get
            {
                foreach(Spreadsheet sheet in Worksheets)
                {
                    if (sheet.isActive())
                        return sheet;
                }
                return null;
            }
        }
        #endregion

        #region Methods
        public void Activate()
        {
            _WBObj.Activate();
        }
        public void Edit(bool show)
        {
            if (show)
            {
                // Check to make sure that the book is not already open
                if (!isOpen)
                {
                    // save + close other open workbooks
                    foreach(Excel.Workbook wb in Control.EditApp.Workbooks)
                    {
                        wb.Save();
                        wb.Close();
                    }
                    _EditObj = Control.EditApp.Workbooks.Open(TempPath);
                }

            }
            else
            {
                _EditObj.Save();
                _EditObj.Close();
                _WBObj.Close();
                _WBObj = Control.App.Workbooks.Open(TempPath, ReadOnly: true);
            }
        }
        public override bool Equals(object obj)
        {
            if(obj.GetType().Equals(typeof(Workbook)))
            {
                return base.Equals(obj);
            }
            else if (obj.GetType().Equals(typeof(Excel.Workbook)))
            {
                return _WBObj.Equals(obj);
            }
            else return base.Equals(obj);
        }
        public void Hide()
        {
            _WinObj.Visible = false;
            isShowing = false;
        }
        public void Show()
        {
            _WinObj.Visible = true;
            isShowing = true;
        }
        public bool isActive()
        {
            return _parent.App.ActiveWorkbook == _WBObj;
        }
        public override string ToString()
        {
            return FileName;
        }

        public void setPath(string path)
        {
            // Parse the path out into it's parts
            Directory = Path.GetDirectoryName(path);
            FileName = Path.GetFileNameWithoutExtension(path);
            FileExtension = Path.GetExtension(path);
            FilePath = Path.GetFullPath(path);
        }
        #endregion
    }
}

