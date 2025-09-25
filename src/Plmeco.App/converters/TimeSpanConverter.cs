using System;
using System.Globalization;
using System.Windows.Data;

namespace Plmeco.App.Converters
{
    public class TimeSpanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is TimeSpan ts) return new DateTime(ts.Ticks).ToString("HH:mm");
            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var s = (value as string)?.Trim() ?? "";
            if (string.IsNullOrEmpty(s)) return null;

            if (TimeSpan.TryParse(s, out var t1)) return new TimeSpan(t1.Hours, t1.Minutes, 0);

            var cleaned = s.Replace(".", ":").Replace(",", ":");
            if (cleaned.Length is 3 or 4) { cleaned = cleaned.PadLeft(4, '0'); cleaned = cleaned.Insert(2, ":"); }
            if (TimeSpan.TryParse(cleaned, out var t2)) return new TimeSpan(t2.Hours, t2.Minutes, 0);

            return Binding.DoNothing;
        }
    }
}
