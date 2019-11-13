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
        #region Fields

        ExcelControl Control
        {
            get
            {
                return Global.Control;
            }
        }
        Workbook ActiveBook
        {
            get
            {
                return Global.ActiveBook;
            }
        }
        Spreadsheet ActiveSheet
        {
            get
            {
                return Global.ActiveSheet;
            }
        }

        #endregion

        #region Constructor
        public pg_ImportWiz()
        {
            InitializeComponent();

            // Initialize UserControls with their event handlers
            tv_sheetList.init(tv_sheetList_SelectedItemChanged, btn_listAdd_Click);

            Unloaded += new RoutedEventHandler(cleanupBeforeClose);
        }
        #endregion

        #region Methods

            #region TreeView
        
        private HierarchicalDataTemplate GetTemplate()
        {
            // Code adapted from:
            //  https://www.codeproject.com/Articles/124644/Basic-Understanding-of-Tree-View-in-WPF#DataBinding

            //create the data template
            HierarchicalDataTemplate dataTemplate = new HierarchicalDataTemplate();

            //create stack pane;
            FrameworkElementFactory grid = new FrameworkElementFactory(typeof(Grid));
            grid.Name = "parentStackpanel";

            // create text
            FrameworkElementFactory label = new FrameworkElementFactory(typeof(TextBlock));
            label.SetBinding(TextBlock.TextProperty, new Binding() { Path = new PropertyPath("Name") });

            grid.AppendChild(label);
            dataTemplate.ItemsSource = new Binding("Worksheets");
            dataTemplate.VisualTree = grid;
            return dataTemplate;
        }
        private void UpdateTreeView()
        {
            TreeView TreeView = tv_sheetList.TreeView;
            TreeView.ItemTemplate = GetTemplate();
            TreeView.ItemsSource = Control.Workbooks;
        }

        private void UnbindTreeView()
        {
            // Unbind TreeView so that it doesn't error when the Control closes
            try
            {
                /*
                Dispatcher.Invoke(() =>
                {
                    TreeView TreeView = tv_sheetList.TreeView;
                    TreeView.ItemsSource = null;
                });
                */
                TreeView TreeView = tv_sheetList.TreeView;
                TreeView.ItemsSource = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        
            #endregion

            #region User Controls
        private void ShowWorkbookControls(bool show)
        {
            if (show)
            {
                // Set UserControl up for ActiveBook

                // Set UserControl Visibility to visible
                uc_Workbook.Visibility = Visibility.Visible;
            }
            else
            {
                // Set UserControl Visibility to collapsed
                uc_Workbook.Visibility = Visibility.Collapsed;
            }
        }
        private void ShowWorksheetControls(bool show)
        {
            if (show)
            {
                // Set UserControl up for ActiveSheet
                uc_Worksheet.populateListView();
                
                // Set UserControl Visibility to visible
                uc_Worksheet.Visibility = Visibility.Visible;
            }
            else
            {
                // Set UserControl Visibility to collapsed
                uc_Worksheet.Visibility = Visibility.Collapsed;
            }
        }

        #endregion

        public void cleanupBeforeClose(object sender, EventArgs e)
        {
            UnbindTreeView();
        }

        #endregion

        #region Event Listeners

            #region Buttons
        private void btn_listAdd_Click(object sender, RoutedEventArgs e)
        {
            // Open up a file dialog for the user to select an Excel (or CSV) document.
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "Excel Workbooks|*.xls;*.xlsx|CSV Files|*.csv|All Files|*.*";
            dlg.ValidateNames = false;
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                try
                {
                    Workbook openBook = Control.addWorkbook(dlg.FileName);
                    openBook.Activate();
                    UpdateTreeView();
                }
                catch (WorkbookLockedException ex)
                {
                    // If the file is locked because it is opened by this program, 
                    //  just select that file from the dropdown.
                    if (!ex.isOpenExternally)
                    {
                        //cb_wbName.SelectedIndex = ex.Index;
                    }
                }
            }
        }

        private void btn_listRem_Click(object sender, RoutedEventArgs e)
        {
            // TODO
        }

        #endregion

            #region TreeView
        private void tv_sheetList_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var obj = tv_sheetList.TreeView.SelectedItem;
            if (obj.GetType().Equals(typeof(Workbook)))
            {
                // Set Active fields to selected item
                //ActiveBook = (Workbook)obj;
                //ActiveSheet = null;
                Workbook book = (Workbook)obj;
                book.Control.App.Visible = true;
                book.Activate();

                // Set textboxes to their correct values
                tb_wbName.Text = ActiveBook.FileName;
                tb_wsName.Text = "";
                tb_wsName.IsEnabled = false;

                // Adjust visibility of User Controls
                ShowWorksheetControls(false);
                ShowWorkbookControls(true);
            }
            else if (obj.GetType().Equals(typeof(Spreadsheet)))
            {
                // Set Active fields to selected item
                //ActiveSheet = (Spreadsheet)obj;
                //ActiveBook = ActiveSheet.Workbook;
                Spreadsheet sheet = (Spreadsheet)obj;
                sheet.Control.App.Visible = true;
                sheet.Activate();

                // Set textboxes to their correct values
                tb_wbName.Text = ActiveBook.FileName;
                tb_wsName.Text = ActiveSheet.Name;
                tb_wsName.IsEnabled = true;

                // Adjust visibility of User Controls
                ShowWorkbookControls(false);
                ShowWorksheetControls(true);
            }
        }
        #endregion

        #endregion
    }
}
