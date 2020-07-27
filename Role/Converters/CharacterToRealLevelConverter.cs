using RoleDDNG.Models.Characters;
using System;
using System.Globalization;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(Personnage), typeof(int))]
    public class CharacterToRealLevelConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Personnage character)
            {
                return character.NiveauGE;
            }
            return 0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int intVal)
            {
                return new Personnage() { NiveauGE = intVal };
            }
            return new Personnage();
        }
    }
}