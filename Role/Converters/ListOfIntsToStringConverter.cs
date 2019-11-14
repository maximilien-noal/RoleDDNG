using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
#pragma warning disable CA1812 // Justification : instantiated from the XAML side

    internal class ListOfIntsToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return "";
            }
            if (value is IList<int> intCollection)
            {
                var stringBuilder = new StringBuilder();
                for (int i = 0; i < intCollection.Count; i++)
                {
                    stringBuilder.Append(intCollection[i].ToString(CultureInfo.CurrentCulture));
                    if (i != intCollection.Count - 1)
                    {
                        stringBuilder.Append(", ");
                    }
                }
                return stringBuilder.ToString();
            }
            return "";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return 0;
            }
            return 0;
        }
    }

#pragma warning restore CA1812
}