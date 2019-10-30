using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    public class Spreadsheet : Excel.WorksheetClass
    {
        private Excel.Worksheet _sheet;
        public Spreadsheet(Excel.Worksheet sheet)
        {
            // TODO
            _sheet = sheet;
            _sheet.CheckSpelling();
        }

        #region Interface Members
        public void Activate()
        {
            _sheet.Activate();
        }
        public void Copy(object Before = null, object After = null)
        {
            if(Equals(Before, null) || Equals(After, null))                // If either before or after are null
            {
                if(Equals(Before, null) && Equals(After, null))                // If BOTH Before AND After are null
                {
                    _sheet.Copy();
                }
                else if(Equals(Before, null))                                  // If just Before is null
                {
                    _sheet.Copy(After: After);
                }
                else if(Equals(After, null))                                  // If just After is null
                {
                    _sheet.Copy(Before: Before);
                }
            }
            else
            {
                _sheet.Copy(Before, After);                             // If neithe Before or After is Null
            }
        }
        public void Calculate()
        {
            _sheet.Calculate();
        }
        #endregion
    }
}
