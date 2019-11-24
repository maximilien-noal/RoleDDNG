﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
#pragma warning disable CA1812 // Justification : instantiated from the XAML side

    internal class IntToStringConverter : IValueConverter
    {
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
            if (value is string str)
            {
                if (int.TryParse(str, out int result))
                {
                    return result;
                }
            }
            return Double.NaN;
        }
    }

#pragma warning restore CA1812
}