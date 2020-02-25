using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(IList<int>), typeof(string))]
    public sealed class ListOfIntsToStringConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        internal static readonly ListOfIntsToStringConverter Default = new ListOfIntsToStringConverter();

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
                return Double.NaN;
            }
            return Double.NaN;
        }
    }
}