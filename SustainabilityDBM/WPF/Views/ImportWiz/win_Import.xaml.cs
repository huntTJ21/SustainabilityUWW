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
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SustainabilityDBM
{
    /// <summary>
    /// Interaction logic for win_Import.xaml
    /// </summary>
    public partial class win_Import : Window
    {
        public win_Import()
        {
            InitializeComponent();

            // Initialize the Global ExelControl
            if (Global.Control == null)
                Global.Control = new ExcelLib.ExcelControl();

            Closing += new CancelEventHandler(cleanupBeforeClose);
        }

        public void cleanupBeforeClose(object sender, EventArgs e)
        {
            //frame.Content = null;
            Global.Control.CleanupAndExit();
        }

        ~win_Import()
        {
            Global.Control.App.Quit();
        }
    }
}
