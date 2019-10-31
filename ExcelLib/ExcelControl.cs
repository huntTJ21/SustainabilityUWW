using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    public class ExcelControl
    {
        public Excel.Application App { get; }
        public List<Workbook> Workbooks { get; }

        public ExcelControl()
        {
            App = new Excel.Application();
            Workbooks = new List<Workbook>();
        }

        ~ExcelControl()
        {
            // Make sure to quit the COM App or it will stay on in the background,
            //  causing file read errors and memory leaks.
            App.Quit();
        }

        // Child Object Constructors
        public void addWorkbook(string fullPath)
        {
            // Make sure the file exists before moving forward
            if (!File.Exists(fullPath)) { throw new FileNotFoundException(); }

            // Create the COM Object, making sure that there are no errors reading the file
            try
            {
                // Create the COM object and store it in a Wrapper Class
                Excel.Workbook WBObj = App.Workbooks.Open(fullPath);
                Workbook newBook = new Workbook(this, WBObj, fullPath);

                // Add the Wraper Object to the list
                Workbooks.Add(newBook);
            }
            catch(System.Runtime.InteropServices.COMException ex)
            {
                // Known Error Codes
                int FileAlreadyOpen = -2146827284;

                if(ex.ErrorCode == FileAlreadyOpen)
                {
                    throw new FileLoadException();
                }
            }
        }
        
    }
}
