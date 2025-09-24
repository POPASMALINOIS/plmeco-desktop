using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using Plmeco.App.Models;
using Plmeco.App.Services;

namespace Plmeco.App
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<LoadRow> Rows { get; } = new();

        public MainWindow()
        {
            InitializeComponent();
            dgCargas.ItemsSource = Rows;
        }

        private void Importar_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xlsx;*.xls|All files|*.*",
                    Title = "Selecciona el fichero de reunión"
                };
                if (dlg.ShowDialog() == true)
                {
                    var data = ImportService.ImportExcel(dlg.FileName);
                    Rows.Clear();
                    foreach (var row in data) Rows.Add(row);
                    dgCargas.Items.Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo importar el archivo:\n" + ex.Message,
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (dgCargas.CurrentItem is not LoadRow row) return;
                if (dgCargas.CurrentColumn is not DataGridColumn col) return;

                var header = col.Header?.ToString() ?? string.Empty;

                // Solo actuamos en las columnas de hora (que en XAML están IsReadOnly)
                if (header.Equals("LLEGADA REAL", StringComparison.OrdinalIgnoreCase))
                {
                    row.LlegadaReal = DateTime.Now.TimeOfDay;
                }
                else if (header.Equals("SALIDA REAL", StringComparison.OrdinalIgnoreCase))
                {
                    row.SalidaReal = DateTime.Now.TimeOfDay;
                }
                else
                {
                    return; // doble clic en otra columna: no hacer nada
                }

                // Refrescar UI sin entrar en modo edición
                dgCargas.CommitEdit(DataGridEditingUnit.Cell, true);
                dgCargas.Items.Refresh();
                e.Handled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al establecer la hora:\n" + ex.Message,
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
