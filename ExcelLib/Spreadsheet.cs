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

        #endregion

        #region Constructor
        public Spreadsheet(Workbook parent, Excel.Worksheet WorksheetObj)
        {
            // Initialize private fields
            _parent = parent;
            _sheetObj = WorksheetObj;
        }
        #endregion

        #region Accessors
        public string Name
        {
            get
            {
                return _sheetObj.Name;
            }
        }
        public ExcelControl ParentApp()
        {
            return _parent.getParentApp();
        }
        public Workbook Workbook
        {
            get
            {
                return _parent;
            }
        }

        #endregion

        #region Methods
        public override string ToString()
        {
            return Name;
        }

        #endregion
    }
}
