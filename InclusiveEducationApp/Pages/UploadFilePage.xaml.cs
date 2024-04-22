using InclusiveEducationApp.Classes;
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
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.DefaultExt = ".xml";
            dialog.Filter = "(.xlxs)|*.xlxs";

            bool? result = dialog.ShowDialog();
                        
            if (result == true)
            {
                // Open document
                string filename = dialog.FileName;
                MessageBox.Show(filename);
            }

            Globals.frame.Navigate(new ActionWithFilePage());
        }
    }
}
