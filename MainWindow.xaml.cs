using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Shapes;
using Path = System.IO.Path;

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

            buttonProcess.IsEnabled =
                !string.IsNullOrEmpty(textBoxSource.Text) &&
                !string.IsNullOrEmpty(textBoxDestination.Text);
        }

        private void ButtonDestination_Click(object sender, RoutedEventArgs e)
        {
            textBoxDestination.Text = OpenFileDialog();

            buttonProcess.IsEnabled =
                !string.IsNullOrEmpty(textBoxSource.Text) &&
                !string.IsNullOrEmpty(textBoxDestination.Text);
        }

        private void ButtonReset_Click(object sender, RoutedEventArgs e)
        {
            textBoxSource.Text = "";
            textBoxDestination.Text = "";
            buttonProcess.IsEnabled = false;
        }

        private async void ButtonProcess_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // copy file to avoid used file
                var sourceFilename = GetTempFilename(textBoxSource.Text);
                var destinationFilename = GetTempFilename(textBoxDestination.Text);

                // load sheets
                var workbook = new XLWorkbook(sourceFilename);
                var source = workbook.Worksheet(1);
                var destination = new XLWorkbook(destinationFilename).Worksheet(1);

                var sepsSource = CellsColumnToList(source, "G", "K", "N");
                var sepsDestination = CellsColumnToList(destination, "F", "J", "I");

                var compare = await CompareAsync(sepsSource, sepsDestination);
                RecolorCell(ref source, compare);

                var filename = SaveFileDialog(textBoxSource.Text);
                if (!string.IsNullOrEmpty(filename))
                {
                    workbook.SaveAs(filename);
                    System.Diagnostics.Process.Start(filename);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Rekap Comparer", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static string OpenFileDialog()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            bool? result = dialog.ShowDialog();
            if (result.HasValue && result.Value)
                return dialog.FileName;

            return null;
        }

        private static string SaveFileDialog(string path)
        {
            var filename = Path.GetFileNameWithoutExtension(path);
            var ext = Path.GetExtension(path);
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                FileName = $"{filename}_Compared{ext}",
            };

            bool? result = dialog.ShowDialog();
            if (result.HasValue && result.Value)
                return dialog.FileName;

            return null;
        }

        private string GetTempFilename(string source)
        {
            var name = Path.GetFileName(source);
            var sourceFilename = Path.Combine(Path.GetTempPath(), name);
            File.Copy(source, sourceFilename, true);
            return sourceFilename;
        }

        private IEnumerable<Rekap> CellsColumnToList(IXLWorksheet worksheet, string sepCol, string trfCol, string gopCol)
        {
            var list = new List<Rekap>();
            var e = 0;
            var i = 1;
            while (e <= 3)
            {
                i++;
                var sep = worksheet?.Cell($"{sepCol}{i}")?.Value?.ToString();
                if (!string.IsNullOrWhiteSpace(sep))
                {
                    var rekap = new Rekap
                    {
                        Id = i,
                        Sep = sep,
                        Tarif = worksheet?.Cell($"{trfCol}{i}")?.Value?.ToString(),
                        Grouping = worksheet?.Cell($"{gopCol}{i}")?.Value?.ToString(),
                    };
                    list.Add(rekap);
                }
                else
                {
                    e++;
                }
            }

            return list;
        }

        private async Task<Dictionary<int, bool>> CompareAsync(IEnumerable<Rekap> source, IEnumerable<Rekap> destination)
        {
            var dictionary = new Dictionary<int, bool>();
            await Task.Run(() =>
            {
                foreach (var rekap in source)
                {
                    var rekap2 = destination.Where(x => x.Sep == rekap.Sep).FirstOrDefault();
                    if (rekap2 != null)
                    {
                        var match = rekap.Tarif == rekap2.Tarif && rekap.Grouping == rekap2.Grouping;
                        dictionary.Add(rekap.Id, match);
                    }
                }
            });
            return dictionary;
        }

        private void RecolorCell(ref IXLWorksheet worksheet, Dictionary<int, bool> rows)
        {
            foreach(var row in rows)
            {
                var color = row.Value ? XLColor.DarkGreen : XLColor.Red;
                worksheet.Range(row.Key, 1, row.Key, 20).Style.Font.FontColor = color;
            }
        }
    }
}
