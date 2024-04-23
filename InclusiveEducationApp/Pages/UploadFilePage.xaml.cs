using InclusiveEducationApp.Classes;
using Microsoft.Win32;
using System;
using System.Windows;
using System.Windows.Controls;

namespace InclusiveEducationApp.Pages
{
    /// <summary>
    /// Логика взаимодействия для UploadFilePage.xaml
    /// </summary>
    public partial class UploadFilePage : Page
    {
        public UploadFilePage()
        {
            InitializeComponent();
        }

        private void upload_xlxs_file_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "Excel Sheet(*.xlsx)|*.xlsx";

            try
            {
                if (op.ShowDialog() == true)
                {
                    Globals.filePath = op.FileName;

                    Globals.frame.Navigate(new ActionWithFilePage());
                }
                else
                {
                    throw new Exception("Не удалось открыть файл");
                }
            }                          
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }                        
        }
    }
}
