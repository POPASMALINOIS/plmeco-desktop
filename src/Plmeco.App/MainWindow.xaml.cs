using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ClosedXML.Excel;
using Microsoft.Win32;
using Plmeco.App.Models;
using Plmeco.App.Services;

namespace Plmeco.App
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<LoadRow> Rows { get; } = new();
        private string? _currentFilePath;

        public MainWindow()
        {
            InitializeComponent();
            dgCargas.ItemsSource = Rows;

            CommandBindings.Add(new CommandBinding(ApplicationCommands.Save, (s,e)=>Guardar(false)));
            CommandBindings.Add(new CommandBinding(ApplicationCommands.SaveAs, (s,e)=>Guardar(true)));
        }

        private void Importar_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls|All files|*.*", Title = "Selecciona el fichero de reunión" };
                if (dlg.ShowDialog() == true)
                {
                    var data = ImportService.ImportExcel(dlg.FileName);
                    Rows.Clear();
                    foreach (var row in data) Rows.Add(row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo importar el archivo:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (dgCargas.CurrentItem is not LoadRow row) return;
                if (dgCargas.CurrentColumn is not DataGridColumn col) return;

                // Finalizar edición antes de tocar datos
                dgCargas.CommitEdit(DataGridEditingUnit.Cell, true);
                dgCargas.CommitEdit(DataGridEditingUnit.Row,  true);
                dgCargas.CancelEdit();

                var header = col.Header?.ToString() ?? string.Empty;
                if (header.Equals("LLEGADA REAL", StringComparison.OrdinalIgnoreCase))
                    row.LlegadaReal = DateTime.Now.TimeOfDay;
                else if (header.Equals("SALIDA REAL", StringComparison.OrdinalIgnoreCase))
                    row.SalidaReal = DateTime.Now.TimeOfDay;
                else
                    return;

                e.Handled = true; // sin Items.Refresh()
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al establecer la hora:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Guardar_Click(object sender, RoutedEventArgs e) => Guardar(false);
        private void GuardarComo_Click(object sender, RoutedEventArgs e) => Guardar(true);
        private void Salir_Click(object sender, RoutedEventArgs e) => Close();

        private void Guardar(bool forcePickPath)
        {
            try
            {
                if (forcePickPath || string.IsNullOrWhiteSpace(_currentFilePath))
                {
                    var sfd = new SaveFileDialog
                    {
                        Title = "Guardar datos",
                        Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv",
                        FileName = "PLMECO_Cargas"
                    };
                    if (sfd.ShowDialog() != true) return;
                    _currentFilePath = sfd.FileName;
                }

                var ext = Path.GetExtension(_currentFilePath!).ToLowerInvariant();
                if (ext == ".csv") ExportToCsv(_currentFilePath!);
                else               ExportToXlsx(_currentFilePath!);

                MessageBox.Show($"Guardado correctamente en:\n{_currentFilePath}", "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo guardar:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportToXlsx(string path)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Cargas");

            var headers = new[] { "MATRICULA","DESTINO","MUELLE","ESTADO","LLEGADA REAL","SALIDA REAL","SALIDA TOPE","INCIDENCIAS","LEX" };
            for (int c = 0; c < headers.Length; c++) ws.Cell(1, c + 1).Value = headers[c];

            int r = 2;
            foreach (var x in Rows)
            {
                ws.Cell(r,1).Value = x.Matricula;
                ws.Cell(r,2).Value = x.Destino;
                ws.Cell(r,3).Value = x.Muelle;
                ws.Cell(r,4).Value = x.Estado;

                if (x.LlegadaReal.HasValue) { ws.Cell(r,5).Value = DateTime.Today + x.LlegadaReal.Value; ws.Cell(r,5).Style.DateFormat.Format = "hh:mm"; }
                if (x.SalidaReal.HasValue)  { ws.Cell(r,6).Value = DateTime.Today + x.SalidaReal.Value;  ws.Cell(r,6).Style.DateFormat.Format = "hh:mm"; }
                if (x.SalidaTope.HasValue)  { ws.Cell(r,7).Value = DateTime.Today + x.SalidaTope.Value;  ws.Cell(r,7).Style.DateFormat.Format = "hh:mm"; }

                ws.Cell(r,8).Value = x.Incidencias;
                ws.Cell(r,9).Value = x.Lex ? "TRUE" : "FALSE";
                r++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
        }

        private void ExportToCsv(string path)
        {
            var sb = new StringBuilder();
            sb.AppendLine("MATRICULA;DESTINO;MUELLE;ESTADO;LLEGADA REAL;SALIDA REAL;SALIDA TOPE;INCIDENCIAS;LEX");

            string Fmt(TimeSpan? t) => t.HasValue ? new DateTime(t.Value.Ticks).ToString("HH:mm") : "";
            string Esc(string? s) => (s ?? string.Empty).Replace("\"", "\"\"");

            foreach (var x in Rows)
            {
                sb.Append('"').Append(Esc(x.Matricula)).Append('"').Append(';')
                  .Append('"').Append(Esc(x.Destino)).Append('"').Append(';')
                  .Append('"').Append(Esc(x.Muelle)).Append('"').Append(';')
                  .Append('"').Append(Esc(x.Estado)).Append('"').Append(';')
                  .Append('"').Append(Fmt(x.LlegadaReal)).Append('"').Append(';')
                  .Append('"').Append(Fmt(x.SalidaReal)).Append('"').Append(';')
                  .Append('"').Append(Fmt(x.SalidaTope)).Append('"').Append(';')
                  .Append('"').Append(Esc(x.Incidencias)).Append('"').Append(';')
                  .Append(x.Lex ? "TRUE" : "FALSE").AppendLine();
            }

            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        }
    }
}
