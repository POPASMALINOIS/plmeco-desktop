using System;
using System.Globalization;
using System.Windows.Data;

namespace Plmeco.App.Converters
{
    // Devuelve True si el texto del binding NO está vacío
    public class NotEmptyConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => !string.IsNullOrWhiteSpace(value?.ToString());

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }
}
