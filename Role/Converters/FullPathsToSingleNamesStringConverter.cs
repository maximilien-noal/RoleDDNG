namespace RoleDDNG.Role.Converters
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Windows.Data;

    [ValueConversion(typeof(IEnumerable<string>), typeof(string))]
    internal sealed class FullPathsToSingleNamesStringConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        private static readonly FullPathsToSingleNamesStringConverter Default = new FullPathsToSingleNamesStringConverter();

        public object? Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is IEnumerable<string> collection)
            {
                var returnValue = "";
                for (int i = 0; i < collection.Count(); i++)
                {
                    var str = collection.ElementAt(i);
                    if (i > 0)
                    {
                        returnValue = $"{returnValue}, ";
                    }
                    if (File.Exists(str))
                    {
                        returnValue = $"{returnValue} {Path.GetFileName(str)}";
                    }
                    else if (Directory.Exists(str))
                    {
                        returnValue = $"{returnValue} {Path.GetDirectoryName(str)}";
                    }
                }
                return returnValue;
            }
            return "";
        }

        public object? ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value?.ToString();
        }
    }
}