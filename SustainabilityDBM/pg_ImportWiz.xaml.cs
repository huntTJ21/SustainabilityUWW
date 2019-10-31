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
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelLib;

namespace SustainabilityDBM
{
    /// <summary>
    /// Interaction logic for pg_ImportWiz.xaml
    /// </summary>
    public partial class pg_ImportWiz : Page
    {
        public pg_ImportWiz()
        {
            InitializeComponent();
        }

        private void btn_fileName_Click(object sender, RoutedEventArgs e)
        {
            // Open up a file dialog for te user to select an Excel (or CSV) document.
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "Excel Workbooks|*.xls;*.xlsx|CSV Files|*.csv|All Files|*.*";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                tb_fileName.Text = dlg.FileName;
                ExcelControl ec = new ExcelControl();
                ec.addWorkbook(@"C:\Users\tjhunt\OneDrive - University of Wisconsin-Whitewater\SAGE\SAGE Member Roster 05-19.xlsx");
                List<Workbook> wb = ec.Workbooks;
                Console.WriteLine("Break");

            }
        }

        private void btn_listAdd_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btn_listRem_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
