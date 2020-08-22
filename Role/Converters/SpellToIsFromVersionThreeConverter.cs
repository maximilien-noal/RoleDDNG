using RoleDDNG.Models.Roles;
using RoleDDNG.Role.Extensions;
using System.Globalization;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(Spell), typeof(bool))]
    public sealed class SpellToIsFromVersionThreeConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        internal static readonly SpellToIsFromVersionThreeConverter Default = new SpellToIsFromVersionThreeConverter();

        public object Convert(object value, System.Type targetType, object parameter, CultureInfo culture)
        {
            return value is Spell spell && spell.IsFromVersion3();
        }

        public object ConvertBack(object value, System.Type targetType, object parameter, CultureInfo culture)
        {
            return new Spell();
        }
    }
}