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

namespace SustainabilityDBM
{
    /// <summary>
    /// Interaction logic for WorkbookTreeView.xaml
    /// </summary>
    public partial class WorkbookTreeView : UserControl
    {
        public WorkbookTreeView()
        {
            InitializeComponent();
        }

        public void init(Action<object, RoutedPropertyChangedEventArgs<object>> SelectedItemChanged, Action<object, RoutedEventArgs> ClickAdd)
        {
            tv.AddHandler(TreeView.SelectedItemChangedEvent, new RoutedPropertyChangedEventHandler<object>(SelectedItemChanged));
            btn_Add.AddHandler(Button.ClickEvent, new RoutedEventHandler(ClickAdd));
        }

        public TreeView TreeView
        {
            get
            {
                return tv;
            }
        }
    }
}
