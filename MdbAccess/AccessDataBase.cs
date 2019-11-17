using System;
using System.Data.Odbc;
using System.IO;

namespace RoleDDNG.MdbAccess
{
    public class AccessDataBase : IDisposable
    {
        private bool _disposed = false;
        private readonly OdbcConnection _dbConnection;

        public AccessDataBase(string mdbFileName)
        {
            if (string.IsNullOrWhiteSpace(mdbFileName))
            {
                throw new ArgumentNullException(nameof(mdbFileName));
            }
            if (File.Exists(mdbFileName) == false)
            {
                throw new FileNotFoundException(mdbFileName);
            }

            _dbConnection = new OdbcConnection("Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + mdbFileName);
            _dbConnection.Open();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }
            if (disposing)
            {
                _dbConnection.Dispose();
            }
            _disposed = true;
        }
    }
}