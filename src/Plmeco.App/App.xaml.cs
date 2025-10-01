using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;

namespace Plmeco.App
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
            base.OnStartup(e);
        }

        private void LogAndShow(string origen, Exception ex)
        {
            try
            {
                var dir = AppDomain.CurrentDomain.BaseDirectory;
                var path = Path.Combine(dir, "plmeco-error.log");
                File.AppendAllText(path, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {origen}: {ex}\r\n\r\n");
            }
            catch { /* ignorar */ }

            MessageBox.Show(
                "Ha ocurrido un error y se ha guardado un log en 'plmeco-error.log'.\n\n" +
                ex.Message,
                "PLMECO",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            LogAndShow("DispatcherUnhandledException", e.Exception);
            e.Handled = true; // <- evita que se cierre la app
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            LogAndShow("UnhandledException", e.ExceptionObject as Exception ?? new Exception("Error no especificado"));
        }

        private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            LogAndShow("UnobservedTaskException", e.Exception);
            e.SetObserved();
        }
    }
}
