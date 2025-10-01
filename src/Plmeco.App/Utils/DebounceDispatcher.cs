using System;
using System.Windows.Threading;

namespace Plmeco.App.Utils
{
    // Llama a la acción solo una vez tras un periodo sin nuevos eventos
    public class DebounceDispatcher
    {
        private DispatcherTimer? _timer;

        public void Debounce(TimeSpan delay, Action action)
        {
            _timer?.Stop();
            _timer = new DispatcherTimer { Interval = delay, Priority = DispatcherPriority.Background };
            _timer.Tick += (s, e) =>
            {
                _timer?.Stop();
                action();
            };
            _timer.Start();
        }
    }
}
