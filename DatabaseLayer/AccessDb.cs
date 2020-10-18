using PetaPoco;

using System.IO;
using System.Threading.Tasks;

namespace RoleDDNG.DatabaseLayer
{
    public sealed class AccessDb
    {
        private const string ConnectionStringBeginning = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";

        private readonly string _dbFilePath = "";

        public AccessDb(string dbFilePath)
        {
            if (string.IsNullOrWhiteSpace(dbFilePath) || File.Exists(dbFilePath) == false)
            {
                throw new FileNotFoundException(dbFilePath);
            }
            _dbFilePath = Path.GetFullPath(dbFilePath);
        }

        public async Task<bool> CanConnectAsync()
        {
            using var database = GetDatabase();
            await database.OpenSharedConnectionAsync().ConfigureAwait(false);
            return database.Connection.State == System.Data.ConnectionState.Open;
        }

        public Database GetDatabase()
        {
            var petaPoco = new Database($"{ConnectionStringBeginning}{_dbFilePath}", "OleDb", new ConventionMapper());
            return petaPoco;
        }
    }
}