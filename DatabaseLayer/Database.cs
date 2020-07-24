using Dapper;
using System.Collections.Generic;
using System.Data.Odbc;
using System.IO;
using System.Threading.Tasks;

namespace RoleDDNG.DatabaseLayer
{
    public class Database
    {
        private const string ConnectionStringBeginning = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=";

        private readonly string _dbFilePath = "";

        public Database(string dbFilePath)
        {
            if (string.IsNullOrWhiteSpace(dbFilePath) || File.Exists(dbFilePath) == false)
            {
                throw new FileNotFoundException(dbFilePath);
            }
            _dbFilePath = Path.GetFullPath(dbFilePath);
        }

        public async Task<bool> CanConnectAsync()
        {
            using OdbcConnection connection = await ConnectAsync().ConfigureAwait(false);
            var connected = connection != null;
            if (connection != null)
            {
                await connection.CloseAsync().ConfigureAwait(false);
            }
            return connected;
        }

        public async Task<IEnumerable<T>> GetQueryDataAsync<T>(string sqlQuery)
        {
            DefaultTypeMap.MatchNamesWithUnderscores = true;
            using OdbcConnection connection = await ConnectAsync().ConfigureAwait(false);
            var result = await connection.QueryAsync<T>(sqlQuery).ConfigureAwait(false);
            await connection.CloseAsync().ConfigureAwait(false);
            return result;
        }

        private async Task<OdbcConnection> ConnectAsync()
        {
            var connection = new OdbcConnection($"{ConnectionStringBeginning}{_dbFilePath}");
            await connection.OpenAsync().ConfigureAwait(false);
            return connection;
        }
    }
}