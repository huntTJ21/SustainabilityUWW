﻿using System;
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
    /// Interaction logic for win_Import.xaml
    /// </summary>
    public partial class win_Import : Window
    {
        public win_Import()
        {
            InitializeComponent();
            
            // Initialize the Global ExelControl
            Global.Control = new ExcelLib.ExcelControl();
        }

        ~win_Import()
        {
            Global.Control.App.Quit();
        }
    }
}
