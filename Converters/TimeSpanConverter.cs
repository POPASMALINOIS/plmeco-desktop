using System;
using System.Globalization;
using System.Windows.Data;

namespace Plmeco.App.Converters
{
    public class TimeSpanConverter : IValueConverter
    {
        // Muestra TimeSpan? como "hh:mm"
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is TimeSpan ts)
                return new DateTime(ts.Ticks).ToString("HH:mm");
            return string.Empty;
        }

        // Acepta texto "hh:mm" (y variantes 730, 7.30, 7,30) -> TimeSpan?
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var s = (value as string)?.Trim() ?? "";
            if (string.IsNullOrEmpty(s)) return null;

            if (TimeSpan.TryParse(s, out var t1))
                return new TimeSpan(t1.Hours, t1.Minutes, 0);

            var cleaned = s.Replace(".", ":").Replace(",", ":");
            if (cleaned.Length == 3 || cleaned.Length == 4) // "730" o "0730"
            {
                cleaned = cleaned.PadLeft(4, '0');
                cleaned = cleaned.Insert(2, ":");
            }
            if (TimeSpan.TryParse(cleaned, out var t2))
                return new TimeSpan(t2.Hours, t2.Minutes, 0);

            // si no se puede, no cambia el valor
            return Binding.DoNothing;
        }
    }
}
