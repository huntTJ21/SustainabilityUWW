using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
//using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelLib;

namespace SustainabilityDBM
{
    /// <summary>
    /// Interaction logic for uc_WSSettings.xaml
    /// </summary>
    public partial class WSSettings : UserControl
    {
        public class lvColor
        {
            public Color Color { get; set; }
            public List<string> Fields
            {
                get
                {
                    return Global.Fields;
                }
            }
            public lvColor(Color c)
            {
                Color = c;
            }

            public override string ToString()
            {
                return Color.ToString();
            }
        }       
        public WSSettings()
        {
            InitializeComponent();
        }

        public void populateListView()
        {
            lv_colors.ItemsSource = Global.ActiveSheet.ColorList;
        }

        private void btn_update_Click(object sender, RoutedEventArgs e)
        {
            Global.ActiveSheet.update();
            populateListView();
        }
    }
}
