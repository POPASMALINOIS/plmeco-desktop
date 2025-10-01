using System;
using System.Collections.Generic;
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
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string n) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));

        public ObservableCollection<DocumentView> Documents { get; } = new();
        private int _selectedIndex;
        public int SelectedIndex { get => _selectedIndex; set { _selectedIndex = value; OnPropertyChanged(nameof(SelectedIndex)); DebouncedAutosave(); } }
        private DocumentView? Current => (SelectedIndex >= 0 && SelectedIndex < Documents.Count) ? Documents[SelectedIndex] : null;

        private readonly DebounceDispatcher _debouncer = new();
        private bool _loadingSnapshot;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            // RESTAURAR
            var snap = PersistenceService.Load();
            _loadingSnapshot = true;
            try
            {
                if (snap.Documents.Count == 0)
                {
                    var doc = new DocumentView { Title = "Reunión 1" };
                    Documents.Add(doc);
                }
                else
                {
                    // ✅ foreach sobre la colección correcta
                    foreach (var d in snap.Documents)
                    {
                        var doc = new DocumentView { Title = d.Title, CurrentFile = d.CurrentFile };
                        foreach (var row in d.Rows) doc.Rows.Add(row);
                        HookRows(doc);
                        Documents.Add(doc);
                    }
                    SelectedIndex = Math.Min(Math.Max(0, snap.SelectedIndex), Documents.Count - 1);
                }
            }
            finally { _loadingSnapshot = false; }

            foreach (var d in Documents) HookRows(d);
            Documents.CollectionChanged += Documents_CollectionChanged;

            SafeAutosaveNow();
        }

        // ===== hooks cambios =====
        private void HookRows(DocumentView doc)
        {
            doc.Rows.CollectionChanged -= Rows_CollectionChanged;
            doc.Rows.CollectionChanged += Rows_CollectionChanged;
            foreach (var r in doc.Rows)
            {
                r.PropertyChanged -= RowOnPropertyChanged;
                r.PropertyChanged += RowOnPropertyChanged;
            }
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
                foreach (var it in e.NewItems.OfType<LoadRow>())
                {
                    it.PropertyChanged -= RowOnPropertyChanged;
                    it.PropertyChanged += RowOnPropertyChanged;
                }
            DebouncedAutosave();
        }
        private void Documents_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (_loadingSnapshot) return;
            DebouncedAutosave();
        }

        private void DebouncedAutosave() => _debouncer.Debounce(TimeSpan.FromSeconds(2), SafeAutosaveNow);

        private void SafeAutosaveNow()
        {
            try
            {
                var docs = new List<PersistenceService.DocumentSnapshot>();
                foreach (var d in Documents)
                {
                    docs.Add(new PersistenceService.DocumentSnapshot
                    {
                        Title = d.Title,
                        CurrentFile = d.CurrentFile,
                        Rows = d.Rows.ToList()
                    });
                }
                // ✅ pasar también SelectedIndex
                PersistenceService.Save(docs, SelectedIndex);
            }
            catch { }
        }

        // ===== importar / pestañas / guardar (igual que envío anterior) =====
        private void ImportarEnActual_Click(object sender, RoutedEventArgs e)
        {
            if (Current is null) return;
            try
            {
                var dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls|All files|*.*", Title = "Selecciona el fichero de reunión" };
                if (dlg.ShowDialog() == true)
                {
                    var data = ImportService.ImportExcel(dlg.FileName);
                    _loadingSnapshot = true;
                    try
                    {
                        Current.Rows.Clear();
                        foreach (var r in data) Current.Rows.Add(r);
                        Current.Title = System.IO.Path.GetFileNameWithoutExtension(dlg.FileName);
                        Current.CurrentFile = null;
                    }
                    finally { _loadingSnapshot = false; }
                    SafeAutosaveNow();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo importar:\n" + ex.Message, "PLMECO", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void NuevaPestanaDesdeExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls|All files|*.*", Title = "Selecciona el fichero de reunión" };
                if (dlg.ShowDialog() == true)
                {
                    var data = ImportService.ImportExcel(dlg.FileName);
                    var title = System.IO.Path.GetFileNameWithoutExtension(dlg.FileName);

                    var doc = new DocumentView { Title = title };
                    foreach (var r in data) doc.Rows.Add(r);
                    HookRows(doc);

                    _loadingSnapshot = true;
                    Documents.Add(doc);
                    SelectedIndex = Documents.Count - 1;
                    _loadingSnapshot = false;

                    SafeAutosaveNow();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo importar en nueva pestaña:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CerrarPestanaActual_Click(object sender, RoutedEventArgs e)
        {
            if (Current is null) return;
            var idx = SelectedIndex;
            if (idx < 0) return;

            if (MessageBox.Show($"¿Cerrar la pestaña \"{Current.Title}\"?",
                                "PLMECO", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            _loadingSnapshot = true;
            Documents.RemoveAt(idx);
            if (Documents.Count == 0)
            {
                var doc = new DocumentView { Title = "Reunión 1" };
                HookRows(doc);
                Documents.Add(doc);
                SelectedIndex = 0;
            }
            else
            {
                SelectedIndex = Math.Min(idx, Documents.Count - 1);
            }
            _loadingSnapshot = false;

            SafeAutosaveNow();
        }

        private void Guardar_Click(object sender, RoutedEventArgs e)
        {
            if (Current is null) return;

            if (string.IsNullOrWhiteSpace(Current.CurrentFile))
            {
                GuardarComo_Click(sender, e);
                return;
            }

            try
            {
                ExportService.ExportExcel(Current.CurrentFile!, Current.Rows);
                MessageBox.Show($"Guardado correctamente:\n{Current.CurrentFile}",
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Information);
                SafeAutosaveNow();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GuardarComo_Click(object sender, RoutedEventArgs e)
        {
            if (Current is null) return;

            var dlg = new SaveFileDialog
            {
                Filter = "Excel Workbook|*.xlsx",
                Title = "Guardar pestaña como",
                FileName = Current.Title.Replace(' ', '_')
            };

            if (dlg.ShowDialog() == true)
            {
                Current.CurrentFile = dlg.FileName;
                Guardar_Click(sender, e);
            }
        }

        private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (sender is not DataGrid grid) return;
                if (grid.CurrentItem is not LoadRow row) return;
                if (grid.CurrentColumn is not DataGridColumn col) return;

                grid.CommitEdit(DataGridEditingUnit.Cell, true);
                grid.CommitEdit(DataGridEditingUnit.Row, true);
                grid.CancelEdit();

                var header = col.Header?.ToString() ?? string.Empty;

                if (header.Equals("LLEGADA REAL", StringComparison.OrdinalIgnoreCase))
                    row.LlegadaReal = DateTime.Now.TimeOfDay;
                else if (header.Equals("SALIDA REAL", StringComparison.OrdinalIgnoreCase))
                    row.SalidaReal = DateTime.Now.TimeOfDay;
                else
                    return;

                e.Handled = true;
                DebouncedAutosave();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al establecer la hora:\n" + ex.Message,
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Salir_Click(object sender, RoutedEventArgs e)
        {
            SafeAutosaveNow();
            Close();
        }
    }
}
