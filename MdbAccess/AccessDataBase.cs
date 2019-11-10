using Dapper;
using System;
using System.Data.Odbc;
using System.Linq;

namespace RoleDDNG.MdbAccess
{
    public class AccessDataBase : IDisposable
    {
        private bool _disposed = false;
        private readonly OdbcConnection _dbConnection;
        private readonly ModelContext _modelContext = new ModelContext();

        public AccessDataBase(string mdbFileName)
        {
            if (string.IsNullOrWhiteSpace(mdbFileName))
            {
                throw new ArgumentNullException(nameof(mdbFileName));
            }

            _dbConnection = new OdbcConnection("Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + mdbFileName);
        }

        public T GetModel<T>() where T : Type
        {
            return _modelContext.Find<T>();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if(_disposed)
            {
                return;
            }
            if(disposing)
            {
                _modelContext.Dispose();
                _dbConnection.Dispose();
            }
            _disposed = true;
        }
    }
}
