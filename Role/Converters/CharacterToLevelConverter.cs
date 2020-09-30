using RoleDDNG.Models.Characters;

using System;
using System.Linq;
using System.Globalization;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(Personnage), typeof(int))]
    public sealed class CharacterToLevelConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        public static readonly CharacterToLevelConverter Default = new CharacterToLevelConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return 0;
            }
            if (value is Personnage character)
            {
                character.EffectiveLevel();
            }
            return 0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return new object();
            }
            return new Personnage();
        }
    }
}