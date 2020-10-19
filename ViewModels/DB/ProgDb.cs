using PetaPoco;

using RoleDDNG.DatabaseLayer;

using System.IO;

namespace RoleDDNG.ViewModels.DB
{
    internal static class ProgDb
    {
        public static Database Create()
        {
            return new AccessDb(Constants.Consts.ProgramDatabaseFileName).GetDatabase();
        }

        public static string GetFullPath()
        {
            return Path.GetFullPath(Constants.Consts.ProgramDatabaseFileName);
        }
    }
}