﻿using System;
using System.Globalization;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(int), typeof(string))]
    public sealed class IntToStringConverter : IValueConverter
    {
        /// <summary>
        /// Gets the default instance
        /// </summary>
        internal static readonly IntToStringConverter Default = new IntToStringConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return "";
            }
            if (value is int integer)
            {
                return integer.ToString(CultureInfo.CurrentCulture);
            }
            return "";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return Double.NaN;
            }
            if (value is string str && int.TryParse(str, out int result))
            {
                return result;
            }
            return Double.NaN;
        }
    }
}