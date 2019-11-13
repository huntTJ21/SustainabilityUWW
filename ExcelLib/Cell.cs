using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    #region Color Classes
    // Color Dictionary
    public static class Colors
    {
        #region Color Fields
        public static Color Black = new Color(0, 0, 0);
        public static Color White = new Color(255, 255, 255);
        public static Color BrightRed = new Color(255, 0, 0);
        public static Color BrightGreen = new Color(255, 0, 0);
        public static Color BrightBlue = new Color(255, 0, 0);
        public static Color BrightYellow = new Color(255, 255, 0);
        public static Color BrightPurple = new Color(255, 0, 255);
        public static Color BrightTeal = new Color(0, 255, 255);
        public static Color DarkRed = new Color(128, 0, 0);
        public static Color DarkGreen = new Color(0, 128, 0);
        public static Color DarkBlue = new Color(0, 0, 128);
        public static Color DarkYellow = new Color(128, 128, 0);
        public static Color DarkPurple = new Color(128, 0, 128);
        public static Color DarkTeal = new Color(0, 128, 128);
        #endregion

        static Dictionary<int, Color> ColorIndexes = new Dictionary<int, Color>
        {
            {1, Black },
            {2, White },
            {3, BrightRed },
            {4, BrightGreen },
            {5, BrightBlue },
            {6, BrightYellow },
            {7, BrightPurple },
            {8, BrightTeal },
            {9, DarkRed },
            {10, DarkGreen },
            {11, DarkBlue },
            {12, DarkYellow },
            {13, DarkPurple },
            {14, DarkTeal },
            {25, DarkBlue },
            {26, BrightPurple },
            {27, BrightYellow },
            {28, BrightTeal },
            {29, DarkPurple },
            {30, DarkRed },
            {31, DarkTeal },
            {32, BrightBlue }
        };
        static void initIndxes()
        {
            // Initialize ColorIndexes
            if (ColorIndexes == null)
            {
                ColorIndexes = new Dictionary<int, Color>();
                ColorIndexes.Add(1, new Color(0, 0, 0));
                ColorIndexes.Add(2, new Color(255, 255, 255));
                ColorIndexes.Add(3, new Color(255, 0, 0));
                ColorIndexes.Add(4, new Color(0, 255, 0));
                ColorIndexes.Add(5, new Color(0, 0, 255));
                ColorIndexes.Add(6, new Color(255, 255, 0));
                ColorIndexes.Add(7, new Color(255, 0, 255));
                ColorIndexes.Add(8, new Color(0, 255, 255));
                ColorIndexes.Add(9, new Color(128, 0, 0));
                ColorIndexes.Add(10, new Color(0, 128, 0));
                ColorIndexes.Add(11, new Color(0, 0, 128));
                ColorIndexes.Add(12, new Color(128, 128, 0));
                ColorIndexes.Add(13, new Color(128, 0, 128));
                ColorIndexes.Add(14, new Color(0, 128, 128));
                ColorIndexes.Add(15, new Color(192, 192, 192));
                ColorIndexes.Add(16, new Color(128, 128, 128));
                ColorIndexes.Add(17, new Color(153, 153, 255));
                ColorIndexes.Add(18, new Color(153, 51, 102));
                ColorIndexes.Add(19, new Color(255, 255, 204));
                ColorIndexes.Add(20, new Color(204, 255, 255));
                ColorIndexes.Add(21, new Color(102, 0, 102));
                ColorIndexes.Add(22, new Color(255, 128, 128));
                ColorIndexes.Add(23, new Color(0, 102, 204));
                ColorIndexes.Add(24, new Color(204, 204, 255));
                ColorIndexes.Add(25, ColorIndexes[11]);
                ColorIndexes.Add(26, ColorIndexes[7]);
                ColorIndexes.Add(27, ColorIndexes[6]);
                ColorIndexes.Add(28, ColorIndexes[8]);
                ColorIndexes.Add(29, ColorIndexes[13]);
                ColorIndexes.Add(30, ColorIndexes[9]);
                ColorIndexes.Add(31, ColorIndexes[14]);
                ColorIndexes.Add(32, ColorIndexes[5]);
                ColorIndexes.Add(33, new Color(0, 204, 255));
                ColorIndexes.Add(34, ColorIndexes[20]);
                ColorIndexes.Add(35, new Color(204, 255, 204));
                ColorIndexes.Add(36, new Color(255, 255, 153));
                ColorIndexes.Add(37, new Color(153, 204, 255));
                ColorIndexes.Add(38, new Color(255, 153, 204));
                ColorIndexes.Add(39, new Color(204, 153, 255));
                ColorIndexes.Add(40, new Color(255, 204, 153));
                ColorIndexes.Add(41, new Color(51, 102, 255));
                ColorIndexes.Add(42, new Color(51, 204, 204));
                ColorIndexes.Add(43, new Color(153, 204, 0));
                ColorIndexes.Add(44, new Color(255, 204, 0));
                ColorIndexes.Add(45, new Color(255, 153, 0));
                ColorIndexes.Add(46, new Color(255, 102, 0));
                ColorIndexes.Add(47, new Color(102, 102, 153));
                ColorIndexes.Add(48, new Color(150, 150, 150));
                ColorIndexes.Add(49, new Color(0, 51, 102));
                ColorIndexes.Add(50, new Color(51, 153, 102));
                ColorIndexes.Add(51, new Color(0, 51, 0));
                ColorIndexes.Add(52, new Color(51, 51, 0));
                ColorIndexes.Add(53, new Color(153, 51, 0));
                ColorIndexes.Add(54, new Color(153, 51, 102));
                ColorIndexes.Add(55, new Color(51, 51, 153));
                ColorIndexes.Add(56, new Color(51, 51, 51));
            }
        }
        
        public static Color ColorFromIndex(int idx)
        {
            if (ColorIndexes.ContainsKey(idx))
                return ColorIndexes[idx];
            else return null;
        }
    }
    
    // Color Class
    public class Color
    {
        #region Fields
        public int R { get; set; }
        public int G { get; set; }
        public int B { get; set; }
        #endregion

        #region Constructors

        public Color(int R, int G, int B)
        {
            this.R = R;
            this.G = G;
            this.B = B;
        }
        
        public Color(double D)
        {
            byte[] b = BitConverter.GetBytes((int)D);
            R = b[0];
            G = b[1];
            B = b[2];
        }

        #endregion

        #region Methods
        public override bool Equals(object obj)
        {
            if(obj.GetType() == typeof(Color))
            {
                Color cObj = (Color)obj;
                if (cObj.R == R && cObj.G == G && cObj.B == B)
                    return true;
                else return false;
            }
            else return base.Equals(obj);
        }
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
        public override string ToString()
        {
            string str = string.Format("({0},{1},{2})", R, G, B);
            return str;
        }
        public double ToDouble()
        {
            byte[] b = new byte[4];
            b[0] = (byte)R;
            b[1] = (byte)G;
            b[2] = (byte)B;
            b[3] = 0;
            int i = BitConverter.ToInt32(b, 0);
            return i;
        }
        public static Color getColor(Excel.Range cell)
        {
            return new Color((double)cell.Interior.Color);
        }
        #endregion
    }
    #endregion

    #region Cell Class
    public class Cell
    {
        #region Fields
        public Excel.Range cell { get; }
        #endregion

        #region Constructor
        public Cell(object rangeObj)
        {
            // Convert object to a range
            Excel.Range range = (Excel.Range)rangeObj;
            
            // Make sure that the given range is only one cell
            if(range.Count != 1)
            {
                throw new Exception("Range is not a single Cell");
            }
            else
            {
                cell = range;
            }
        }
        #endregion

        #region Accessors
        public int ColorIndex
        {
            get { return (int)cell.Interior.ColorIndex;}
        }

        public Color Color
        {
            get 
            {
                return new Color((double)cell.Interior.Color);
            }
            set
            {
                cell.Interior.Color = value.ToDouble();
            }
        }

        public string Value
        {
            get
            {
                return cell.Value.ToString();
            }
            set
            {
                cell.Value = value;
            }
        }
        #endregion

        #region Methods

        #endregion

    }
    #endregion
}

