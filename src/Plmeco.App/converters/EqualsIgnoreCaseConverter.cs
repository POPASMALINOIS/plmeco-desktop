using System;
using System.Globalization;
using System.Windows.Data;

namespace Plmeco.App.Converters
{
    // Devuelve True si el texto del binding == ConverterParameter (ignorando mayúsc/minúsc)
    public class EqualsIgnoreCaseConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var s = value?.ToString() ?? "";
            var p = parameter?.ToString() ?? "";
            return string.Equals(s, p, StringComparison.OrdinalIgnoreCase);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }
}
