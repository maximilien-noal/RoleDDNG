using RoleDDNG.DatabaseLayer.Enums;

namespace RoleDDNG.DatabaseLayer.Models
{
    public class Database
    {
        public Database(string filePath, DbType dbType)
        {
            DbType = dbType;
            FilePath = filePath;
        }

        public DbType DbType { get; set; } = DbType.UserCharactersDb;

        public string FilePath { get; set; } = "";
    }
}