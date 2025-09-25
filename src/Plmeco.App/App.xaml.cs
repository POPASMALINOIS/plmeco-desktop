using System.Windows;

namespace Plmeco.App
{
    public partial class App : Application
    {
        public App()
        {
            // Captura cualquier excepción no controlada del hilo de UI
            this.DispatcherUnhandledException += (s, e) =>
            {
                MessageBox.Show(
                    "Ha ocurrido un error:\n" + e.Exception.Message,
                    "PLMECO",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
                e.Handled = true; // Evita que la app se cierre
            };
        }
    }
}
