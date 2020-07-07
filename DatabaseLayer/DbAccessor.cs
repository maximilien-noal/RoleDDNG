using Dapper;

using RoleDDNG.DatabaseLayer.Models;

using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Threading.Tasks;

namespace RoleDDNG.DatabaseLayer
{
    public sealed class DbAccessor
    {
        private const string ConnectionStringBeginning = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=";

        private readonly Database _database;

        public DbAccessor(Database db)
        {
            _database = db;
        }

        public async Task<IEnumerable<T>> GetQueryDataAsync<T>(string sqlQuery)
        {
            DefaultTypeMap.MatchNamesWithUnderscores = true;
            using var connection = new OdbcConnection($"{ConnectionStringBeginning}{_database.FilePath}");
            await connection.OpenAsync().ConfigureAwait(false);
            var result = connection.Query<T>(sqlQuery);
            await connection.CloseAsync().ConfigureAwait(false);
            return result;
        }
    }
}