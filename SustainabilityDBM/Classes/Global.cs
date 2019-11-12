using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SustainabilityDBM
{
    public static class Global
    {
        #region ExcelLib Objects
        public static ExcelLib.ExcelControl Control { get; set; }
        public static ExcelLib.Workbook ActiveBook
        {
            get
            {
                return Control.ActiveBook;
            }
        }
        public static ExcelLib.Spreadsheet ActiveSheet
        {
            get
            {
                return Control.ActiveSheet;
            }
        }
        #endregion

        #region Database Objects
        public static List<string> Fields { get; set; }
        #endregion
    }
}
