using System.Windows;

namespace Plmeco.App
{
    public partial class App : Application
    {
        public App()
        {
            this.DispatcherUnhandledException += (s, e) =>
            {
                MessageBox.Show("Ha ocurrido un error: " + e.Exception.Message,
                                "PLMECO", MessageBoxButton.OK, MessageBoxImage.Error);
                e.Handled = true; // evita que la app se cierre
            };
        }
    }
}
