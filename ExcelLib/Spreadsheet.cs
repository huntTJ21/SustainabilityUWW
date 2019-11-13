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
        public Excel.Worksheet _sheetObj;
        #endregion

        #region Public
        public Dictionary<Color, List<int>> ColorDict = new Dictionary<Color, List<int>>();
        public List<Color> ColorList { get; private set; }
        #endregion

        #endregion

        #region Constructor
        public Spreadsheet(Workbook parent, Excel.Worksheet WorksheetObj)
        {
            // Initialize private fields
            _parent = parent;
            _sheetObj = WorksheetObj;

            // Do the initial update
            update();
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
        public ExcelControl Control
        {
            get
            {
                return _parent.Control;
            }
        }
        public Workbook Workbook
        {
            get
            {
                return _parent;
            }
        }
        public Cell this[int row, int col]
        {
            get
            {
                return new Cell((Excel.Range)_sheetObj.Cells[row, col]);
            }
        }
        public Cell[,] this[int r1, int c1, int r2, int c2]
        {
            get
            {
                // Make it so that r1 and c1 are the larger of the two numbers
                if (r2 > r1)
                {
                    int t = r1;
                    r1 = r2;
                    r2 = t;
                }
                if (c2 > c1)
                {
                    int t = c1;
                    c1 = c2;
                    c2 = t;
                }
                // Get number of rows and cols               
                int n_rows = r1 - r2;
                int n_cols = r1 - r2;

                // Calculate size of Cell array
                Cell[,] cells = new Cell[n_rows, n_cols];

                // Retrieve cell objects and put them into the Cell array
                for (int r = r1; r <= r2; r++)
                {
                    for (int c = c1; c <= c2; c++)
                    {
                        // Calculate array index
                        cells[r - r1, c - c1] = this[r, c];
                    }
                }

                // Return the Cell array
                return cells;
            }
        }
        #endregion

        #region Methods
        public void Activate()
        {
            _sheetObj.Activate();
        }
        public bool isActive()
        {
            return _sheetObj == _parent.Control.App.ActiveSheet;
        }

        public void update()
        {
            // Create new Color list
            List<Color> cl = new List<Color>();

            // Go through each cell in the UsedRange
            foreach (Excel.Range cell in _sheetObj.UsedRange)
            {
                // Check if each cell has a background color
                if ((double)cell.Interior.Color != Colors.White.ToDouble())
                {
                    // Check if the color has already been added to the ColorList
                    Color c = Color.getColor(cell);
                    if (!cl.Contains(c))
                    {
                        // Add it to the ColorList
                        cl.Add(c);
                    }
                }
            }

            // Update the ColorList
            ColorList = cl;

        }
        public override string ToString()
        {
            return Name;
        }

        #endregion
    }
}