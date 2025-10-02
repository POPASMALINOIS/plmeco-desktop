using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Plmeco.App.Converters
{
    public class TimeSpanTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) return string.Empty;
            if (value is TimeSpan ts)
                return ts.Hours.ToString("00") + ":" + ts.Minutes.ToString("00");
            if (value is DateTime dt)
                return dt.Hour.ToString("00") + ":" + dt.Minute.ToString("00");
            return value.ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var s = (value == null ? string.Empty : value.ToString().Trim());
            if (s.Length == 0) return null;

            s = s.Replace(".", ":").Replace(",", ":");

            if (TimeSpan.TryParse(s, CultureInfo.InvariantCulture, out var t1))
                return new TimeSpan(t1.Hours, t1.Minutes, 0);

            var digits = new string(s.Where(char.IsDigit).ToArray());
            if (digits.Length == 3 || digits.Length == 4)
            {
                if (digits.Length == 3) digits = "0" + digits;
                var fixedStr = digits.Substring(0, 2) + ":" + digits.Substring(2, 2);
                if (TimeSpan.TryParse(fixedStr, CultureInfo.InvariantCulture, out var t2))
                    return new TimeSpan(t2.Hours, t2.Minutes, 0);
            }
            return Binding.DoNothing;
        }
    }
}
