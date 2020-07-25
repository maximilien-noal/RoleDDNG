using Dapper;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.IO;
using System.Threading.Tasks;

namespace RoleDDNG.DatabaseLayer
{
    public sealed class Database : IDisposable, IAsyncDisposable
    {
        private const string ConnectionStringBeginning = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=";

        private readonly string _dbFilePath = "";

        private OdbcConnection? _connection = null;

        private bool _wasDisposedAlready;

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
            if (_connection is null)
            {
                _connection = await ConnectAsync().ConfigureAwait(false);
            }
            return _connection is null == false;
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        public ValueTask DisposeAsync()
        {
            Dispose();
            return default;
        }

        public async Task<IEnumerable<T>> GetQueryDataAsync<T>(string sqlQuery)
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

        private void Dispose(bool disposing)
        {
            if (!_wasDisposedAlready)
            {
                if (disposing)
                {
                    _connection?.Dispose();
                }
                _wasDisposedAlready = true;
            }
        }
    }
}