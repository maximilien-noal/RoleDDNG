namespace RoleDDNG.ViewModels.DB
{
    internal static class CommonQueries
    {
        public static string DbCharactersAll { get; private set; } = "select nom,image,race,niv_1,niv_2,niv_3,niv_4,niv_5,niv_6,niv_7,niv_8,classe_1,classe_2,classe_3,classe_4,classe_5,classe_6,classe_7,classe_8 from personnage where exclu=false order by nom";
        public static string DbCharactersNames { get; private set; } = "select nom from personnage where exclu=false";
    }
}