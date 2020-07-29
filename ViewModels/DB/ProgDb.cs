using PetaPoco;
using RoleDDNG.DatabaseLayer;

namespace RoleDDNG.ViewModels.DB
{
    internal static class ProgDb
    {
        public static Database Create()
        {
            return new AccessDb(Constants.Consts.ProgramDatabaseFileName).GetDatabase();
        }
    }
}