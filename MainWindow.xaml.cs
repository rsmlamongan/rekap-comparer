using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace RekapComparer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ButtonSource_Click(object sender, RoutedEventArgs e)
        {
            textBoxSource.Text = OpenFileDialog();
        }

        private void ButtonDestination_Click(object sender, RoutedEventArgs e)
        {
            textBoxDestination.Text = OpenFileDialog();
        }

        private void ButtonProcess_Click(object sender, RoutedEventArgs e)
        {

        }

        private static string OpenFileDialog()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx"
            };

            bool? result = dialog.ShowDialog();
            if (result.HasValue && result.Value)
                return dialog.FileName;

            return null;
        }
    }
}
