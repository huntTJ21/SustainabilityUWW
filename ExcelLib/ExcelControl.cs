using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    public class ExcelControl
    {
        #region Fields
        public Excel.Application App { get; }
        public List<Workbook> Workbooks { get; }
        #endregion

        #region Constructor and Destructor
        public ExcelControl()
        {
            // Initialize members
            App = new Excel.Application();
            Workbooks = new List<Workbook>();

            // Set up temp directory
            if (Directory.Exists(@".\ExcelTemp"))
            {
                Directory.Delete(@".\ExcelTemp", true);
            }
            Directory.CreateDirectory(@".\ExcelTemp");
        }

        ~ExcelControl()
        {
            // Make sure to quit the COM App or it will stay on in the background,
            //  causing file read errors and memory leaks.
            App.Quit();

            // Make sure to clear temp directory
            Directory.Delete(@".\ExcelTemp", true);
        }
        #endregion

        #region Child Object Constructors
        public void addWorkbook(string fullPath)
        {
            // Make sure the file exists before moving forward
            if (!File.Exists(fullPath)) { throw new FileNotFoundException(); }

            // Make sure that the file is not already opened
            for(int i = 0; i < Workbooks.Count; i++)
            {
                // If the file being attempted to open is already opened by this instance, 
                //  throw a WorkbookLockedException with the index of the opened workbook.
                if(Workbooks[i].FilePath == Path.GetFullPath(fullPath))
                {
                    throw new WorkbookLockedException(this, i);
                }
            }

            // Make a temporary copy of the workbook, to avoid file locks
            string tempPath = createTempWBFile(fullPath);

            // Create the COM Object, making sure that there are no errors reading the file
            Excel.Workbook WBObj = null;
            try
            {
                // Create the COM object and store it in a Wrapper Class
                WBObj = App.Workbooks.Open(fullPath);
                Workbook newBook = new Workbook(this, WBObj, tempPath);             // Load the workbook from the temp file to avoid file locks
                newBook.setPath(fullPath);                                          // set the path to the original path instead of the temp path
                newBook.TempPath = Path.GetFullPath(tempPath);                      // Set the tempPath variable for future reference.
                
                // Add the Wraper Object to the list
                Workbooks.Add(newBook);
            }
            catch(System.Runtime.InteropServices.COMException ex)
            {
                // Known Error Codes
                int FileAlreadyOpen = -2146827284;

                // If the workbook opened, make sure it is closed to prevent memory leaks
                if (WBObj != null)
                {
                    WBObj.Close();
                }

                if(ex.ErrorCode == FileAlreadyOpen)
                {
                    throw new WorkbookLockedException(true);
                }
            }
        }
        #endregion

        #region Methods

        private string createTempWBFile(string fullPath)
        {
            string tempPath = @".\ExcelTemp";
            string fileName = Path.GetFileNameWithoutExtension(fullPath);
            string fileExt = Path.GetExtension(fullPath);
            fileName += "_temp" + fileExt;
            string newPath = Path.Combine(tempPath, fileName);
            File.Copy(fullPath, newPath);
            return newPath;
        }

        #endregion
    }

    [Serializable]
    public class WorkbookLockedException : Exception
    {
        public string ResourceReferenceProperty { get; set; }
        public ExcelControl ExcelControl { get; }
        public Workbook Workbook { get; }
        public int Index { get; }
        public WorkbookLockedException(){}
        public bool isOpenExternally { get; }
        public WorkbookLockedException(bool isOpenExternally) { this.isOpenExternally = isOpenExternally; }
        public WorkbookLockedException(string msg, bool isOpenExternally) : base(msg) { this.isOpenExternally = isOpenExternally; }
        public WorkbookLockedException(ExcelControl activeEC, Workbook openWB) {
            ExcelControl = activeEC;
            Workbook = openWB;
            isOpenExternally = false;
        }
        public WorkbookLockedException(ExcelControl ec, int wbIndex)
        {
            ExcelControl = ec;
            Index = wbIndex;
        }
        public WorkbookLockedException(string msg, Exception inner) : base(msg, inner) { }
        protected WorkbookLockedException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
            ResourceReferenceProperty = info.GetString("ResourceReferenceProperty");
        }

        [SecurityPermission(SecurityAction.Demand, SerializationFormatter = true)]
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            if (info == null)
                throw new ArgumentNullException("info");
            info.AddValue("ResourceReferenceProperty", ResourceReferenceProperty);
            base.GetObjectData(info, context);
        }
    }
}
