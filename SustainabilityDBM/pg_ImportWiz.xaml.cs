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
        ExcelControl ec;
        public pg_ImportWiz()
        {
            InitializeComponent();
            ec = new ExcelControl();
        }

        private Workbook[] LoadComboBoxData()
        {
            Workbook[] wbRange = new Workbook[ec.Workbooks.Count];
            for(int i = 0; i < ec.Workbooks.Count; i++)
            {
                wbRange[i] = ec.Workbooks[i];
            }
            return wbRange;
        }


        private void btn_wbName_Click(object sender, RoutedEventArgs e)
        {
            // Open up a file dialog for te user to select an Excel (or CSV) document.
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "Excel Workbooks|*.xls;*.xlsx|CSV Files|*.csv|All Files|*.*";
            dlg.ValidateNames = false;
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                try
                {
                    ec.addWorkbook(dlg.FileName);
                    cb_wbName.ItemsSource = LoadComboBoxData();
                    cb_wbName.SelectedIndex = ec.Workbooks.Count - 1;
                }
                catch (WorkbookLockedException ex)
                {
                    // If the file is locked because it is opened by this program, 
                    //  just select that file from the dropdown.
                    if (!ex.isOpenExternally)
                    {
                        cb_wbName.SelectedIndex = ex.Index;
                    }
                }
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
