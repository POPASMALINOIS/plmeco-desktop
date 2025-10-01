using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using PLMECO.App.Models;
using PLMECO.App.Services;

namespace PLMECO.App
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<DocumentView> Documents { get; set; } = new ObservableCollection<DocumentView>();
        private bool _loadingSnapshot = false;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            // Crear pestaña vacía inicial
            var doc = new DocumentView { Title = "Hoja 1" };
            Documents.Add(doc);
            HookRows(doc);
        }

        // === PROPIEDADES DE AYUDA ===
        public int SelectedIndex
        {
            get => TabControlDocs.SelectedIndex;
            set => TabControlDocs.SelectedIndex = value;
        }

        public DocumentView Current => (DocumentView)TabControlDocs.SelectedItem;

        // === IMPORTAR EN PESTAÑA ACTUAL ===
        private void ImportarEnActual_Click(object sender, RoutedEventArgs e)
        {
            if (Current == null) return;
            try
            {
                var dlg = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xlsx;*.xls|All files|*.*",
                    Title = "Selecciona el fichero de reunión"
                };

                if (dlg.ShowDialog() == true)
                {
                    var res = ImportService.ImportExcel(dlg.FileName);

                    if (res.Rows.Count == 0)
                    {
                        MessageBox.Show(
                            "No se han encontrado filas importables.\n\n" +
                            "Cabeceras mínimas: MATRICULA, DESTINO, MUELLE, ESTADO.\n" +
                            "Comprueba también que no haya filas 'LADO/SECCIÓN/CARGA' antes de la fila de títulos.",
                            "PLMECO", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    _loadingSnapshot = true;
                    try
                    {
                        Current.Rows.Clear();
                        foreach (var r in res.Rows) Current.Rows.Add(r);
                        Current.Title = res.SheetName;     // nombre real de la hoja
                        Current.CurrentFile = null;
                        HookRows(Current);
                    }
                    finally { _loadingSnapshot = false; }

                    SafeAutosaveNow();

                    MessageBox.Show($"Importado desde hoja \"{res.SheetName}\" (fila cabecera {res.HeaderRow}).\n" +
                                    $"Filas: {res.Rows.Count}.",
                                    "PLMECO", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo importar:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // === IMPORTAR EN NUEVA PESTAÑA ===
        private void NuevaPestanaDesdeExcel_Click(object sender, RoutedEventArgs e)
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
                    var res = ImportService.ImportExcel(dlg.FileName);

                    if (res.Rows.Count == 0)
                    {
                        MessageBox.Show(
                            "No se han encontrado filas importables.\n\n" +
                            "Cabeceras mínimas: MATRICULA, DESTINO, MUELLE, ESTADO.",
                            "PLMECO", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    var doc = new DocumentView { Title = res.SheetName };
                    foreach (var r in res.Rows) doc.Rows.Add(r);
                    HookRows(doc);

                    _loadingSnapshot = true;
                    Documents.Add(doc);
                    SelectedIndex = Documents.Count - 1;  // selecciona la nueva pestaña
                    _loadingSnapshot = false;

                    SafeAutosaveNow();

                    MessageBox.Show($"Importado en nueva pestaña desde \"{res.SheetName}\".\n" +
                                    $"Filas: {res.Rows.Count}.",
                                    "PLMECO", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo importar en nueva pestaña:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // === GUARDAR Y GUARDAR COMO ===
        private void Guardar_Click(object sender, RoutedEventArgs e)
        {
            if (Current == null) return;
            try
            {
                if (string.IsNullOrWhiteSpace(Current.CurrentFile))
                    GuardarComo_Click(sender, e);
                else
                {
                    ExportService.ExportToExcel(Current.CurrentFile, Current.Rows);
                    MessageBox.Show("Guardado correctamente en:\n" + Current.CurrentFile,
                                    "PLMECO", MessageBoxButton.OK, MessageBoxImage.Information);
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
                var dlg = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Guardar pestaña como..."
                };

                if (dlg.ShowDialog() == true)
                {
                    Current.CurrentFile = dlg.FileName;
                    ExportService.ExportToExcel(dlg.FileName, Current.Rows);
                    MessageBox.Show("Guardado como:\n" + dlg.FileName,
                                    "PLMECO", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar como:\n" + ex.Message, "PLMECO",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // === CERRAR PESTAÑA ===
        private void CerrarPestana_Click(object sender, RoutedEventArgs e)
        {
            if (Current != null)
                Documents.Remove(Current);
        }

        // === AUTOSAVE PARA NO PERDER DATOS ===
        private void SafeAutosaveNow()
        {
            if (_loadingSnapshot) return;
            try
            {
                PersistenceService.SaveSnapshot(Documents.ToList());
            }
            catch { /* no interrumpir */ }
        }

        private void HookRows(DocumentView doc)
        {
            doc.Rows.CollectionChanged += (s, e) => SafeAutosaveNow();
            foreach (var row in doc.Rows)
                row.PropertyChanged += (s, e) => SafeAutosaveNow();
        }
    }
}
