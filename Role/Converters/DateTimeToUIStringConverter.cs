using System;
using System.Globalization;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(DateTime), typeof(string))]
    public sealed class DateTimeToUIStringConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        public static readonly DateTimeToUIStringConverter Default = new DateTimeToUIStringConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is DateTime dt)
            {
                return dt.ToString("g", CultureInfo.CurrentCulture);
            }
            return "";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return DateTime.Now;
        }
    }
}