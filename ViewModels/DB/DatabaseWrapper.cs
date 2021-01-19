using GalaSoft.MvvmLight.Ioc;
using RoleDDNG.DatabaseLayer;
using RoleDDNG.Models.Options;
using PetaPoco;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;

namespace RoleDDNG.ViewModels.DB
{
    internal static class DatabaseWrapper
    {
        public static Database CreateProgDb()
        {
            return new AccessDb(Constants.Consts.ProgramDatabaseFileName).GetDatabase();
        }

        public static Database CreateCharactersDb(string source = "")
        {
            if (string.IsNullOrWhiteSpace(source))
            {
                source = SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath;
            }
            return new AccessDb(source).GetDatabase();
        }

        public static async Task<C> GetCollectionFromQueryAsync<T, C>(string query, string source = "") where C : IList<T>, new()
        {
            var collection = new C();
            using var charactersDb = CreateCharactersDb(source);
            using var elementsReader = await charactersDb.QueryAsync<T>(query).ConfigureAwait(true);
            while (await elementsReader.ReadAsync().ConfigureAwait(true))
            {
                collection.Add(elementsReader.Poco);
            }
            return collection;
        }

        public static string GetFullProgDbPath() => Path.GetFullPath(Constants.Consts.ProgramDatabaseFileName);
    }
}