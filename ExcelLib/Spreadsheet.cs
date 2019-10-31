using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelLib
{
    public class Spreadsheet
    {
        #region Fields

        #region Private
        private Workbook _parent;
        private Excel.Worksheet _sheetObj;
        #endregion

        #region Public
        #endregion

        public Spreadsheet(Workbook parent, Excel.Worksheet WorksheetObj)
        {
            // Initialize private fields
            _parent = parent;
            _sheetObj = WorksheetObj;
        }

        #region Accessors
        public string SheetName()
        {
            return _sheetObj.Name;
        }
        public ExcelControl ParentApp()
        {
            return _parent.getParentApp();
        }

        public Workbook Workbook()
        {
            return _parent;
        }

        #endregion

        #endregion
    }
}
