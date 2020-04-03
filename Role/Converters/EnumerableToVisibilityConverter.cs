using System;
using System.Collections;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Linq;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(IEnumerable), typeof(Visibility))]
    public sealed class EnumerableToVisibilityConverter : IValueConverter
    {
        /// <summary>
        /// Gets the default instance
        /// </summary>
        public static readonly EnumerableToVisibilityConverter Default = new EnumerableToVisibilityConverter();

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
            var count = ((IEnumerable)value).OfType<object>().Count();
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