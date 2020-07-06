using Dapper;

using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Threading.Tasks;

namespace RoleDDNG.DatabaseLayer
{
    public sealed class DbAccessor : IDbAccessor, IDisposable
    {
        private const string ConnectionStringBeginning = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=";

        private readonly string _dbFileName;

        private OdbcConnection? _connection;

        private bool isDisposed;

        public DbAccessor(string filename)
        {
            _dbFileName = filename;
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        public async Task<IEnumerable<T>> GetQueryDataAsync<T>(string sqlQuery)
        {
            DefaultTypeMap.MatchNamesWithUnderscores = true;
            if(_connection is null)
            {
                _connection = new OdbcConnection($"{ConnectionStringBeginning}{_dbFileName}");
            }
            await _connection.OpenAsync().ConfigureAwait(false);
            var result = _connection.Query<T>(sqlQuery);
            await _connection.CloseAsync().ConfigureAwait(false);
            return result;
        }

        private void Dispose(bool disposing)
        {
            if (!isDisposed)
            {
                if (disposing && _connection != null)
                {
                    _connection.Dispose();
                }
                isDisposed = true;
            }
        }
    }
}