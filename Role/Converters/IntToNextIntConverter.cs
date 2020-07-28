using System;
using System.Globalization;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(int), typeof(int))]
    public sealed class IntToNextIntConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        internal static readonly IntToNextIntConverter Default = new IntToNextIntConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int intValue && intValue < Int32.MaxValue)
            {
                return intValue + 1;
            }
            return 0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int intValue && intValue > Int32.MinValue)
            {
                return intValue - 1;
            }
            return 0;
        }
    }
}