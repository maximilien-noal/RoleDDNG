using RoleDDNG.Models.Roles;
using System;

namespace RoleDDNG.Role.Extensions
{
    internal static class SpellExt
    {
        public static bool IsEpic(this Spell spell)
        {
            return string.IsNullOrWhiteSpace(spell.ClasseNiveau) == false && spell.ClasseNiveau.ToUpperInvariant().Contains("EPIQUE", StringComparison.InvariantCultureIgnoreCase);
        }

        public static bool IsFromVersion3(this Spell spell)
        {
            return spell.Version != "3.5";
        }
    }
}