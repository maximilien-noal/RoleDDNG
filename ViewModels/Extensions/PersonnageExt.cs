using RoleDDNG.Models.Characters;

namespace RoleDDNG.ViewModels.Extensions
{
    internal static class PersonnageExt
    {
        public static double GetXpPointsForLevel(this Personnage character)
        {
            return GetXpPointsForLevel(character.NiveauGE);
        }

        public static double GetXpPointsForNextLevel(this Personnage character)
        {
            return GetXpPointsForLevel(character.NiveauGE + 1);
        }

        private static double GetXpPointsForLevel(double level)
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