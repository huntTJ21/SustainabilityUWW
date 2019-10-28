using System;
using System.Collections.Generic;
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
    /// Interaction logic for win_MainMenu.xaml
    /// </summary>
    public partial class win_MainMenu : Window
    {
        public win_MainMenu()
        {
            InitializeComponent();
        }

        private void btn_Settings_Click(object sender, RoutedEventArgs e)
        {
            var settingsWin = new win_Settings { WindowStartupLocation = this.WindowStartupLocation };
            settingsWin.Closing += delegate { this.Show(); };
            settingsWin.Show();
            this.Hide();
        }
    }
}
