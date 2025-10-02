using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using Plmeco.App.Models;
using Plmeco.App.Services;

namespace Plmeco.App
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<DocumentView> Documents { get; } = new();
        public int SelectedIndex
        {
            get => TabControlDocs.SelectedIndex;
            set => TabControlDocs.SelectedIndex = value;
        }
        public DocumentView Current => (DocumentView)TabControlDocs.SelectedItem;

        // lista para el Combo de ESTADO
        public System.Collections.Generic.List<string> EstadosDisponibles { get; } =
            new() { "CARGANDO", "OK", "ANULADO" };

        private bool _loadingSnapshot;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            var doc = new DocumentView { Title = "Hoja 1" };
            Documents.Add(doc);
            HookRows(doc);
        }

        private void HookRows(DocumentView doc)
        {
            doc.Rows.CollectionChanged += (s, e) => SafeAutosaveNow();
            foreach (var r in doc.Rows)
                r.PropertyChanged += (s, e) => SafeAutosaveNow();
        }

        private void SafeAutosaveNow()
        {
            if (_loadingSnapshot) return;
            try { PersistenceService.SaveSnapshot(Documents.ToList()); } catch { }
        }

        private void ImportarEnActual_Click(object sender, RoutedEventArgs e)
        {
            if (Current == null) return;
            try
            {
                var dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls|All files|*.*", Title = "Selecciona el fichero de reunión" };
                if (dlg.ShowDialog() != true) return;

                var res = ImportService.ImportExcel(dlg.FileName);
                if (res.Rows.Count == 0)
                {
                    MessageBox.Show("No se han encontrado filas importables.\n" +
                                    "Necesarias: MATRICULA, DESTINO, MUELLE, ESTADO.",
                                    "PLMECO", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _loadingSnapshot = true;
                try
                {
                    Current.Rows.Clear();
                    foreach (var r in res.Rows) Current.Rows.Add(r);
                    Current.Title = res.SheetName;
                    Current.CurrentFile = null;
                    HookRows(Current);
                }
                finally { _loadingSnapshot = false; }

                SafeAutosaveNow();
                MessageBox.Show($"Importado \"{res.SheetName}\" (cabecera fila {res.HeaderRow}) — Filas: {res.Rows.Count}.",
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo importar:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void NuevaPestanaDesdeExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls|All files|*.*", Title = "Selecciona el fichero de reunión" };
                if (dlg.ShowDialog() != true) return;

                var res = ImportService.ImportExcel(dlg.FileName);
                if (res.Rows.Count == 0)
                {
                    MessageBox.Show("No se han encontrado filas importables.\n" +
                                    "Necesarias: MATRICULA, DESTINO, MUELLE, ESTADO.",
                                    "PLMECO", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var doc = new DocumentView { Title = res.SheetName };
                foreach (var r in res.Rows) doc.Rows.Add(r);
                HookRows(doc);

                _loadingSnapshot = true;
                Documents.Add(doc);
                SelectedIndex = Documents.Count - 1;
                _loadingSnapshot = false;

                SafeAutosaveNow();
                MessageBox.Show($"Importado en nueva pestaña \"{res.SheetName}\" — Filas: {res.Rows.Count}.",
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo importar en nueva pestaña:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Guardar_Click(object sender, RoutedEventArgs e)
        {
            if (Current == null) return;
            try
            {
                if (string.IsNullOrWhiteSpace(Current.CurrentFile))
                    GuardarComo_Click(sender, e);
                else
                {
                    ExportService.ExportExcel(Current.CurrentFile, Current.Rows);
                    MessageBox.Show("Guardado en:\n" + Current.CurrentFile, "PLMECO",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GuardarComo_Click(object sender, RoutedEventArgs e)
        {
            if (Current == null) return;
            try
            {
                var dlg = new SaveFileDialog { Filter = "Excel Files|*.xlsx", Title = "Guardar pestaña como..." };
                if (dlg.ShowDialog() == true)
                {
                    Current.CurrentFile = dlg.FileName;
                    ExportService.ExportExcel(dlg.FileName, Current.Rows);
                    MessageBox.Show("Guardado como:\n" + dlg.FileName, "PLMECO",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar como:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CerrarPestanaActual_Click(object sender, RoutedEventArgs e)
        {
            if (Current != null) Documents.Remove(Current);
        }
    }
}
