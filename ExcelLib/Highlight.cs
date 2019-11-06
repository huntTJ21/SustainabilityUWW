using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLib
{
    class Highlight
    {
        public class Color
        {
            int R {get; set;}
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


    }
}
