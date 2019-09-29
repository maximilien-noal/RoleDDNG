using System;
using Xunit;

namespace MdbAccess.Tests
{
    public class AccessDataBaseTests
    {
        [Fact]
        public void TestConnection()
        {
            Assert.True(AccessDataBase.TestConnection());
        }
    }
}
