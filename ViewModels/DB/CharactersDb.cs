using GalaSoft.MvvmLight.Ioc;
using RoleDDNG.DatabaseLayer;
using RoleDDNG.Models.Options;
using PetaPoco;

namespace RoleDDNG.ViewModels.DB
{
    internal static class CharactersDb
    {
        public static Database Create(string source = "")
        {
            if (string.IsNullOrWhiteSpace(source))
            {
                source = SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath;
            }
            return new AccessDb(source).GetDatabase();
        }
    }
}