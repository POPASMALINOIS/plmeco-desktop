using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using Plmeco.App.Models;
using Plmeco.App.Services;
using Plmeco.App.Utils;

namespace Plmeco.App
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<LoadRow> Rows { get; } = new();
        private string? currentFile; // último archivo de exportación
        private readonly DebounceDispatcher _debouncer = new();
        private bool _loadingSnapshot = false;

        public MainWindow()
        {
            InitializeComponent();

            // Restaurar snapshot si existe
            var snapshot = PersistenceService.Load();
            _loadingSnapshot = true;
            try
            {
                Rows.Clear();
                foreach (var r in snapshot) Rows.Add(r);
            }
            finally
            {
                _loadingSnapshot = false;
            }

            dgCargas.ItemsSource = Rows;

            // Suscribir cambios para autosave
            Rows.CollectionChanged += Rows_CollectionChanged;
            foreach (var row in Rows) HookRow(row);

            // Opcional: primer guardado al inicializar (para crear el fichero si no existe)
            SafeAutosaveNow();
        }

        // ====== AUTOSAVE ======
        private void HookRow(LoadRow row)
        {
            row.PropertyChanged -= RowOnPropertyChanged; // evitar doble suscripción
            row.PropertyChanged += RowOnPropertyChanged;
        }

        private void RowOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (_loadingSnapshot) return;
            DebouncedAutosave();
        }

        private void Rows_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (_loadingSnapshot) return;

            if (e.NewItems != null)
                foreach (var item in e.NewItems.OfType<LoadRow>())
                    HookRow(item);

            DebouncedAutosave();
        }

        private void DebouncedAutosave()
        {
            _debouncer.Debounce(TimeSpan.FromSeconds(2), SafeAutosaveNow);
        }

        private void SafeAutosaveNow()
        {
            try { PersistenceService.Save(Rows); }
            catch { /* silencioso */ }
        }

        // ====== IMPORTAR ======
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
                    _loadingSnapshot = true;
                    try
                    {
                        Rows.Clear();
                        foreach (var row in data) Rows.Add(row);
                    }
                    finally
                    {
                        _loadingSnapshot = false;
                    }
                    SafeAutosaveNow(); // snapshot inmediato tras importar
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo importar el archivo:\n" + ex.Message,
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ====== DOBLE CLIC HORAS ======
        private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (dgCargas.CurrentItem is not LoadRow row) return;
                if (dgCargas.CurrentColumn is not DataGridColumn col) return;

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
                DebouncedAutosave(); // guarda tras el cambio
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al establecer la hora:\n" + ex.Message,
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ====== GUARDAR / GUARDAR COMO (exporta a Excel con colores) ======
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
                // No borramos la copia de seguridad: así, si mañana se abre la app, todavía está el último snapshot.
                // Si quieres borrarla al exportar, descomenta la línea siguiente:
                // PersistenceService.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar:\n" + ex.Message,
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

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
            // Guardado rápido antes de salir
            SafeAutosaveNow();
            this.Close();
        }
    }
}
