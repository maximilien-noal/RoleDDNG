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

#pragma warning disable CA1031 // Do not catch general exception types

        public async Task<bool> CanConnectAsync()
        {
            try
            {
                using OdbcConnection connection = await ConnectAsync().ConfigureAwait(false);
                await connection.CloseAsync().ConfigureAwait(false);
                return true;
            }
            catch (Exception)
            {
            }
            return false;
        }

#pragma warning restore CA1031 // Do not catch general exception types

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
            var connection = new OdbcConnection($"{ConnectionStringBeginning}{_database.FilePath}");
            await connection.OpenAsync().ConfigureAwait(false);
            return connection;
        }
    }
}