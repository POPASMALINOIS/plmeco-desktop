using System;
using System.Windows.Threading;

namespace Plmeco.App.Utils
{
    // Ejecuta la acción una única vez tras un periodo sin nuevos eventos.
    public class DebounceDispatcher
    {
        private DispatcherTimer? _timer;

        public void Debounce(TimeSpan delay, Action action)
        {
            _timer?.Stop();

            // En WPF se usa el constructor con prioridad (no existe propiedad Priority)
            _timer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = delay
            };

            EventHandler? handler = null;
            handler = (s, e) =>
            {
                _timer!.Stop();
                _timer.Tick -= handler!;
                action();
            };
            _timer.Tick += handler;
            _timer.Start();
        }
    }
}
