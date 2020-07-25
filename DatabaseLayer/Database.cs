using Dapper;
using System.Collections.Generic;
using System.Data.Odbc;
using System.IO;
using System.Threading.Tasks;

namespace RoleDDNG.DatabaseLayer
{
    public sealed class Database
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
            DefaultTypeMap.MatchNamesWithUnderscores = true;
        }

        public async Task<bool> CanConnectAsync()
        {
            using var connection = await ConnectAsync().ConfigureAwait(false);
            return connection is null == false;
        }

        public async Task<IEnumerable<T>> QuerySingleAsync<T>(string sqlQuery)
        {
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