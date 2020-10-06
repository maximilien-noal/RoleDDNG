using RoleDDNG.Models.Roles;

using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(string), typeof(bool))]
    public sealed class StringToBooleanConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        internal static readonly StringToBooleanConverter Default = new StringToBooleanConverter();

        public object Convert(object value, System.Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null || value is string str && string.IsNullOrWhiteSpace(str))
            {
                return false;
            }
            return true;
        }

        public object ConvertBack(object value, System.Type targetType, object parameter, CultureInfo culture)
        {
            return "";
        }
    }
}