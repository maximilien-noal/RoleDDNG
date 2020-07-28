using RoleDDNG.Models.Characters;

using System;
using System.Globalization;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(Personnage), typeof(bool))]
    public sealed class CharacterToIsNextLevelConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        public static readonly CharacterToIsNextLevelConverter Default = new CharacterToIsNextLevelConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return 0;
            }
            if (value is Personnage character)
            {
                return character.Xp >= character.NiveauSuivant;
            }
            return false;
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