using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Plmeco.App.Converters
{
    public class TimeSpanTextConverter : IValueConverter
    {
        // object -> string (muestra)
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is TimeSpan ts)
                return $"{ts.Hours:00}:{ts.Minutes:00}";
            if (value is TimeSpan? nts && nts.HasValue)
                return $"{nts.Value.Hours:00}:{nts.Value.Minutes:00}";
            if (value is DateTime dt)
                return $"{dt.Hour:00}:{dt.Minute:00}";
            return ""; // vacío si null o no válido
        }

        // string -> object (edición)
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var s = (value?.ToString() ?? "").Trim();
            if (string.IsNullOrEmpty(s)) return null!;

            // normaliza separadores y casos como "730" -> "07:30"
            s = s.Replace(".", ":").Replace(",", ":");
            if (TimeSpan.TryParse(s, CultureInfo.InvariantCulture, out var t1))
                return new TimeSpan(t1.Hours, t1.Minutes, 0);

            var digits = new string(s.Where(char.IsDigit).ToArray());
            if (digits.Length is 3 or 4)
            {
                digits = digits.PadLeft(4, '0');
                var fixedStr = digits.Insert(2, ":");
                if (TimeSpan.TryParse(fixedStr, CultureInfo.InvariantCulture, out var t2))
                    return new TimeSpan(t2.Hours, t2.Minutes, 0);
            }

            // si no se puede convertir, no cambies el modelo:
            return Binding.DoNothing;
        }
    }
}
