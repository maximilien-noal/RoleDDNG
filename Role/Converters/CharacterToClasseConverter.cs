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
                var levels = new short?[] { character.Niv1, character.Niv2, character.Niv3, character.Niv4, character.Niv5, character.Niv6, character.Niv7, character.Niv8 };
                var classes = new string[] { character.Classe1, character.Classe2, character.Classe3, character.Classe4, character.Classe5, character.Classe6, character.Classe7, character.Classe8 };
                var maxLevel = levels.Max();
                var topClasse = classes[levels.ToList().IndexOf(maxLevel)];
                return topClasse;
            }
            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null)
            {
                return null;
            }
            return new Personnage();
        }
    }
}