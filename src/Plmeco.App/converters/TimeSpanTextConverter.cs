using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Plmeco.App.Converters
{
    /// <summary>
    /// Conversor robusto para mostrar/editar horas en formato HH:mm.
    /// - Convert: object -> string "HH:mm"
    /// - ConvertBack: string (07:30, 7.30, 730, 0730) -> TimeSpan
    /// Nunca lanza excepciones; si no puede convertir, devuelve Binding.DoNothing.
    /// </summary>
    public class TimeSpanTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) return string.Empty;

            // TimeSpan
            if (value is TimeSpan)
            {
                var ts = (TimeSpan)value;
                return ts.Hours.ToString("00") + ":" + ts.Minutes.ToString("00");
            }

            // DateTime
            if (value is DateTime)
            {
                var dt = (DateTime)value;
                return dt.Hour.ToString("00") + ":" + dt.Minute.ToString("00");
            }

            // Cualquier otro tipo: devuelve ToString() tal cual
            return value.ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var s = (value == null ? string.Empty : value.ToString().Trim());
            if (s.Length == 0) return null;

            // Normaliza separadores
            s = s.Replace(".", ":").Replace(",", ":");

            // Intento directo
            TimeSpan t;
            if (TimeSpan.TryParse(s, CultureInfo.InvariantCulture, out t))
                return new TimeSpan(t.Hours, t.Minutes, 0);

            // “730” / “0730” -> “07:30”
            var digits = new string(s.Where(char.IsDigit).ToArray());
            if (digits.Length == 3 || digits.Length == 4)
            {
                if (digits.Length == 3) digits = "0" + digits;
                var fixedStr = digits.Substring(0, 2) + ":" + digits.Substring(2, 2);
                if (TimeSpan.TryParse(fixedStr, CultureInfo.InvariantCulture, out t))
                    return new TimeSpan(t.Hours, t.Minutes, 0);
            }

            // No tocar el modelo si no se puede convertir
            return Binding.DoNothing;
        }
    }
}
