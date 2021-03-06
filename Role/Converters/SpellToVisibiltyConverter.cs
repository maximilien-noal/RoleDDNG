﻿using RoleDDNG.Models.Roles;

using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace RoleDDNG.Role.Converters
{
    [ValueConversion(typeof(Spell), typeof(Visibility))]
    public sealed class SpellToVisibiltyConverter : IValueConverter
    {
        /// <summary> Gets the default instance </summary>
        internal static readonly SpellToVisibiltyConverter Default = new SpellToVisibiltyConverter();

        public object Convert(object value, System.Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Spell spell && spell.IsEpic())
            {
                return Visibility.Visible;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, System.Type targetType, object parameter, CultureInfo culture)
        {
            return new Spell();
        }
    }
}