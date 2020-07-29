using GalaSoft.MvvmLight.Ioc;
using RoleDDNG.DatabaseLayer;
using RoleDDNG.Models.Options;
using PetaPoco;

namespace RoleDDNG.ViewModels.DB
{
    internal static class CharactersDb
    {
        public static Database Create()
        {
            return new AccessDb(SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath).GetDatabase();
        }
    }
}