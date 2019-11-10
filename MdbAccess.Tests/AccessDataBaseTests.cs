using Dapper;
using RoleDDNG.MdbAccess;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace MdbAccess.Tests
{
    public class AccessDataBaseTests
    {
        //public static IEnumerable<object[]> GetConnection()
        //{
        //    //Construct the query
        //    //Usually you use @parameter as the syntax for querying SQL Server
        //    //But for querying to Microsoft Access you must use ?
        //    var queryText = "SELECT * FROM Personnage WHERE Nom = ?";

        //    //Use dapper to query with parameter
        //    //It's also a good idea if you use a string as a parameter that you use DbString instead of sending the variable directly
        //    //You will also need to specify the Length exactly as the length of the column in the Microsoft Access table
        //    var data = connection.Query<object>(queryText, new { Name = new DbString { Value = "Rebelle", Length = 0, IsAnsi = true } });
        //}

        //[Fact]
        //public void TestConnection()
        //{
        //    Assert.True(AccessDataBase.TestConnection());
        //}
    }
}
