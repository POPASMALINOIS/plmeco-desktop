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
        private string? currentFile; // Para recordar el último archivo guardado

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

                // Finalizar edición antes de modificar datos
                dgCargas.CommitEdit(DataGridEditingUnit.Cell, true);
                dgCargas.CommitEdit(DataGridEditingUnit.Row, true);
                dgCargas.CancelEdit();

                var header = col.Header?.ToString() ?? string.Empty;

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
                    return;
                }

                e.Handled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al establecer la hora:\n" + ex.Message,
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // --- GUARDAR ---
        private void Guardar_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(currentFile))
            {
                GuardarComo_Click(sender, e);
                return;
            }

            try
            {
                ExportService.ExportExcel(currentFile, Rows);
                MessageBox.Show("Archivo guardado correctamente",
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar:\n" + ex.Message,
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // --- GUARDAR COMO ---
        private void GuardarComo_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Excel Workbook|*.xlsx",
                Title = "Guardar como"
            };

            if (dlg.ShowDialog() == true)
            {
                currentFile = dlg.FileName;
                Guardar_Click(sender, e);
            }
        }

        private void Salir_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
