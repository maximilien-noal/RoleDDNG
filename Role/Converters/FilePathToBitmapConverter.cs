using System;
using System.Globalization;
using System.IO;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(string), typeof(BitmapImage))]
    public sealed class FilePathToBitmapConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        internal static readonly FilePathToBitmapConverter Default = new FilePathToBitmapConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return new object();
            }
            if (value is string filePath && string.IsNullOrWhiteSpace(filePath) == false && File.Exists(filePath))
            {
                var bitmapRender = new BitmapImage();
                bitmapRender.BeginInit();
                bitmapRender.StreamSource = File.OpenRead(filePath);
                bitmapRender.EndInit();
                return bitmapRender;
            }
            return new object();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return string.Empty;
            }
            if (value is BitmapImage img)
            {
                return img.ToString(CultureInfo.InvariantCulture);
            }
            return string.Empty;
        }
    }
}