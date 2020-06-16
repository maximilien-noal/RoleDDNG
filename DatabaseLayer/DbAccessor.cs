using Dapper;

using System.Collections.Generic;
using System.Data.Odbc;
using System.Threading.Tasks;

namespace RoleDDNG.DatabaseLayer
{
    public sealed class DbAccessor : IDbAccessor
    {
        private const string ConnectionStringBeginning = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=";

        private readonly string _dbFileName;

        public DbAccessor(string filename)
        {
            _dbFileName = filename;
        }

        public async Task<IEnumerable<T>> GetQueryDataAsync<T>(string sqlQuery)
        {
            DefaultTypeMap.MatchNamesWithUnderscores = true;
            using var connection = new OdbcConnection($"{ConnectionStringBeginning}{_dbFileName}");
            await connection.OpenAsync().ConfigureAwait(false);
            var result = connection.Query<T>(sqlQuery);
            await connection.CloseAsync().ConfigureAwait(false);
            return result;
        }
    }
}