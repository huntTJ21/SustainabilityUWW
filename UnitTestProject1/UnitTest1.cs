using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelLib;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1() 
        {
            ExcelControl ec = new ExcelControl();
            ec.addWorkbook(@"C:\Users\tjhunt\OneDrive - University of Wisconsin-Whitewater\SAGE\SAGE Member Roster 05-19.xlsx");
            object wb = ec.Workbooks;
            Console.WriteLine("Break");
        }
    }
}
