﻿using System;
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
        #endregion

        #region Public
        public List<Spreadsheet> Worksheets { get; }
        public string Directory { get; set; }
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string FileExtension { get; set; }
        public string TempPath { get; set; }
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
            _WBObj = parent.App.Workbooks.Open(fullPath);
            _parent = parent;

            Worksheets = new List<Spreadsheet>();

            // Populate Spreadsheet List
            foreach (Excel.Worksheet sheet in _WBObj.Worksheets)
            {
                Worksheets.Add(new Spreadsheet(this, sheet));
            }

            // Add event handler
            //Control.App.WorkbookBeforeClose += WorkbookBeforeCloseHandler;
            void WorkbookBeforeCloseHandler(Excel.Workbook Wb, ref bool Cancel)
            {
                //Cancel = true;
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

