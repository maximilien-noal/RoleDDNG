using Dapper;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;

namespace MdbAccess
{
    public static class AccessDataBase
    {
        public static bool TestConnection()
        {
            //Create a connection
            using (var connection = new OdbcConnection(@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=..\..\..\..\MdbAccess.Tests\MDBs\perso.mdb"))
            {

                //Construct the query
                //Usually you use @parameter as the syntax for querying SQL Server
                //But for querying to Microsoft Access you must use ?
                var queryText = "SELECT * FROM Personnage WHERE Nom = ?";

                //Use dapper to query with parameter
                //It's also a good idea if you use a string as a parameter that you use DbString instead of sending the variable directly
                //You will also need to specify the Length exactly as the length of the column in the Microsoft Access table
                var data = connection.Query<object>(queryText, new { Name = new DbString { Value = "Rebelle", Length = 50, IsAnsi = true } });
                bool canRetrivedData = data.Any();
                return canRetrivedData;
            }
        }
    }
}
