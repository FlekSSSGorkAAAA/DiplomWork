﻿using InclusiveEducationApp.Classes;
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

namespace InclusiveEducationApp.Pages
{
    /// <summary>
    /// Логика взаимодействия для ActionWithFilePage.xaml
    /// </summary>
    public partial class ActionWithFilePage : Page
    {
        public ActionWithFilePage()
        {
            InitializeComponent();
        }

        private void makeCountOperation_Click(object sender, RoutedEventArgs e)
        {

        }

        private void makeLoadInfo_Click(object sender, RoutedEventArgs e)
        {
            Globals.frame.Navigate(new LoadInfoToExcelPage());
        }
    }
}
