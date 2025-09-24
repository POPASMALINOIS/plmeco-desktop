using System;
using System.Collections.ObjectModel;
using System.Windows;
using Microsoft.Win32;
using Plmeco.App.Models;
using Plmeco.App.Services;

namespace Plmeco.App
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<LoadRow> Rows { get; set; } = new();

        public MainWindow()
        {
            InitializeComponent();
            dgCargas.ItemsSource = Rows;
        }

        private void Importar_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls" };
            if (dlg.ShowDialog() == true)
            {
                var data = ImportService.ImportExcel(dlg.FileName);
                Rows.Clear();
                foreach (var row in data) Rows.Add(row);
            }
        }

        private void DataGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (dgCargas.CurrentItem is LoadRow row)
            {
                var col = dgCargas.CurrentColumn.Header.ToString();
                if (col == "LLEGADA REAL") row.LlegadaReal = DateTime.Now.TimeOfDay;
                if (col == "SALIDA REAL") row.SalidaReal = DateTime.Now.TimeOfDay;
                dgCargas.Items.Refresh();
            }
        }
    }
}
