namespace RoleDDNG.ViewModels.DB
{
    internal static class CommonQueries
    {
        public const string DbCharactersAll = "select * from personnage where exclu=false order by nom";
        public const string DbCharactersQuick = "select nom,image,race,niv_1,niv_2,niv_3,niv_4,niv_5,niv_6,niv_7,niv_8,classe_1,classe_2,classe_3,classe_4,classe_5,classe_6,classe_7,classe_8 from personnage where exclu=false order by nom";
        public const string DbCharactersNames = "select nom from personnage where exclu=false";
        public const string GetObjetsNames = "select NomObjet from Objets";
        public const string GetAllObjects = "select * from Objets";
        public const string GetAllObjectsPropriete = "select * from ObjetsPropriete";
    }
}