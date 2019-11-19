using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

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
        public static List<ExcelLib.Color> ActiveColorList
        {
            get
            {
                return ActiveSheet.ColorList;
            }
        }
        #endregion

        #region Database Objects
        public static List<string> DBFields { get; set; }
        #endregion

        #region WPF Properties
        public static List<dynamic> WorksheetListView
        {
            get
            {
                List<dynamic> list = new List<dynamic>();
                foreach(ExcelLib.Color c in ActiveColorList)
                {
                    System.Windows.Media.Color mC = Color.FromRgb(
                        Convert.ToByte(c.R.ToString()),
                        Convert.ToByte(c.G.ToString()),
                        Convert.ToByte(c.B.ToString())
                        );

                    SolidColorBrush b = new SolidColorBrush(mC);
                    var obj = new { Text = c.ToString(),  Brush = b, Fields = DBFields, Sheet = ActiveSheet};
                    list.Add(obj);
                }
                return list;
            }
        }
        #endregion

        public static void init()
        {
            Control = null;
        }

        public static void HideActiveWindow()
        {
            Control.App.Windows[1].Visible = false;
        }
        public static void ShowActiveWindow()
        {
            Control.App.Windows[1].Visible = true;
        }
    }
}
