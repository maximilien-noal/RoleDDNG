using RoleDDNG.Models.Characters;

namespace RoleDDNG.ViewModels.Extensions
{
    internal static class PersonnageExt
    {
        public static long GetNextLevelXp(this Personnage character)
        {
            return GetXpLevel(character.NiveauGE + 1);
        }

        public static long GetXpLevel(this Personnage character)
        {
            return GetXpLevel(character.NiveauGE);
        }

        private static long GetXpLevel(long level)
        {
            var xpLevel = 0;

            if (level < 2)
            {
                return xpLevel;
            }

            var i = 1;
            while (i < level)
            {
                xpLevel += (i * 1000);
                i++;
            }

            return xpLevel;
        }
    }
}