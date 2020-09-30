using RoleDDNG.Models.Characters;

using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(Personnage), typeof(string))]
    public sealed class CharacterToClasseConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        public static readonly CharacterToClasseConverter Default = new CharacterToClasseConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return string.Empty;
            }
            if (value is Personnage character)
            {
                return character.EffectiveClass();
            }
            return string.Empty;
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