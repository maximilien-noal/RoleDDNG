using System;
using System.Collections;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    public class EnumerableToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return Visibility.Collapsed;
            }
            if (value is IEnumerable == false)
            {
                return Visibility.Visible;
            }
            var count = 0;
            foreach (var item in (IEnumerable)value)
            {
                count = 1;
                break;
            }
            if (count > 0)
            {
                return Visibility.Visible;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Array.Empty<object>();
        }
    }
}