using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    enum test
    {
        None=0,
        Black,
        White,
        Red,


    }
    
    public class Highlight
    {
        #region Color Class
        public class Color
        {
            int R { get; set; }
            int G { get; set; }
            int B { get; set; }
            public Color(int R, int G, int B)
            {
                this.R = R;
                this.G = G;
                this.B = B;
            }

            public override string ToString()
            {
                string str = String.Format("({0},{1},{2})", R, G, B);
                return base.ToString();
            }
        }
        #endregion

        #region Color Dicitonaries
        public static Dictionary<int, Color> ColorIndexes = new Dictionary<int, Color>()
        {
            { 1, new Color(0,0,0) },
            { 2, new Color(255, 255, 255) },
            { 3, new Color(255, 0, 0) },
            { 4, new Color(0, 255 ,0) },
            { 5, new Color(0, 0, 255) },
            { 6, new Color(255, 255, 0) },
            { 7, new Color(255, 0, 255) },
            { 8, new Color(0, 255, 255) },
            { 9, new Color(128, 0, 0) },
            { 10, new Color(0, 128, 0) },
            { 11, new Color(0, 0, 128) },
            { 12, new Color(128, 128, 0) },
            { 13, new Color(128, 0, 128) },
            { 14, new Color(0, 128, 128) },
            { 15, new Color(192, 192, 192) },
            { 16, new Color(128, 128, 128) },
            { 17, new Color(153, 153, 255) },
            { 18, new Color(153, 51, 102) },
            { 19, new Color(255, 255, 204) },
            { 20, new Color(204, 255, 255) },
            { 21, new Color(102, 0, 102) },
            { 22, new Color(255, 128, 128) },
            { 23, new Color(0, 102, 204) },
            { 24, new Color(204, 204, 255) },
            { 25, ColorIndexes[11] },
            { 26, ColorIndexes[7] },
            { 27, ColorIndexes[6] },
            { 28, ColorIndexes[8] },
            { 29, ColorIndexes[13] },
            { 30,  ColorIndexes[9] },
            { 31, ColorIndexes[14] },
            { 32, ColorIndexes[5] },
            { 33, new Color(0, 204, 255) },
            { 34, ColorIndexes[20] },
            { 35, new Color(204, 255, 204) },
            { 36, new Color(255, 255, 153) },
            { 37, new Color(153, 204, 255) },
            { 38, new Color(255, 153, 204) },
            { 39, new Color(204, 153, 255) },
            { 40, new Color(255, 204, 153) },
            { 41, new Color(51, 102, 255) },
            { 42, new Color(51, 204, 204) },
            { 43, new Color(153, 204, 0) },
            { 44, new Color(255, 204, 0) },
            { 45, new Color(255, 153, 0) },
            { 46, new Color(255, 102, 0) },
            { 47, new Color(102, 102, 153) },
            { 48, new Color(150, 150, 150) },
            { 49, new Color(0, 51, 102) },
            { 50, new Color(51, 153, 102) },
            { 51, new Color(0, 51, 0) },
            { 52, new Color(51, 51, 0) },
            { 53, new Color(153, 51, 0) },
            { 54, new Color(153, 51, 102) },
            { 55, new Color(51, 51, 153) },
            { 56, new Color(51, 51, 51) }
        };

        public Dictionary<string, Color> ColorNames = new Dictionary<string, Color>()
        {
            {"Black", ColorIndexes[1] },
            {"White", ColorIndexes[2] },
            {"Bright Red", ColorIndexes[3] }
        };
        #endregion

        #region testFuncions
        public static Excel.Worksheet blankSheet()
        {
            ExcelControl ec = new ExcelControl();
            Excel.Workbook book = ec.App.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)book.ActiveSheet;
            return sheet;
        }
        public static void testHighlights(object obj)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)obj;
            sheet.Application.Visible = true;
            for (int i = 0; i < 7; i++)
            {
                for (int j = 0; j < 8; j++)
                {
                    int index = (i * 8) + j;
                    Excel.Range cell = (Excel.Range)sheet.Cells[i + 1, j + 1];
                    cell.Interior.ColorIndex = index;
                    cell.Value = "" + index;
                }
            }
        }
        #endregion
    }

    public class Cell
    {
        public Excel.Range cell { get; }

        public Cell(object rangeObj)
        {
            Excel.Range range = (Excel.Range)rangeObj;
            if(range.Count != 1)
            {
                throw new Exception('Range is not a single Cell');
            }
            else
            {
                cell = range;
            }
        }

        public int ColorIndex
        {
            get { return (int)cell.Interior.ColorIndex;}
        }

        public Highlight.Color Color
        {
            get { return Highlight.ColorIndexes[ColorIndex]; }
        }
    }
}
