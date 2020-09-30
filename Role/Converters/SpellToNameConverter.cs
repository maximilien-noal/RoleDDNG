using RoleDDNG.Models.Roles;

using System.Globalization;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(Spell), typeof(string))]
    public sealed class SpellToNameConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        internal static readonly SpellToNameConverter Default = new SpellToNameConverter();

        public object Convert(object value, System.Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Spell spell)
            {
                if (spell.IsFromVersion3())
                {
                    return $"{spell.Nom} (version 3.0)";
                }
                return spell.Nom is null ? "" : spell.Nom;
            }
            return "";
        }

        public object ConvertBack(object value, System.Type targetType, object parameter, CultureInfo culture)
        {
            return new Spell();
        }
    }
}