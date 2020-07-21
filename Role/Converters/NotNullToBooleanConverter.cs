using System;
using System.Globalization;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(object), typeof(bool))]
    public sealed class NotNullToBooleanConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        internal static readonly NotNullToBooleanConverter Default = new NotNullToBooleanConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return false;
            }
            return true;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null || value is bool boolValue && boolValue == false)
            {
                return false;
            }
            return true;
        }
    }
}