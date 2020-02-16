using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.IO;
using System.Threading.Tasks;

using RoleDDNG.Interfaces.Repository;

namespace RoleDDNG.DataAccess
{
    public class AccessDataBase<T> : IAsyncRepository<T> where T : new()
    {
        public async Task<IEnumerable<T>> GetAllAsync(string mdbFilePath, string tableName)
        {
            if (string.IsNullOrWhiteSpace(mdbFilePath))
            {
                throw new ArgumentNullException(nameof(mdbFilePath));
            }
            if (File.Exists(mdbFilePath) == false)
            {
                throw new FileNotFoundException(mdbFilePath);
            }
            using var db = new OdbcConnection("Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + mdbFilePath);
            await db.OpenAsync().ConfigureAwait(false);

            var query = $"SELECT * FROM {tableName}";

            return new List<T>();
        }
    }
}