using PetaPoco;

using System.Threading.Tasks;

using RoleDDNG.Models.Characters;
using System.Linq;

namespace RoleDDNG.ViewModels.DB
{
    internal class MigrationsRunner
    {
        private bool _doneMigrations;

        private const string DiceRollMigration =
            "CREATE TABLE DiceRoll (" +
            "Id INTEGER PRIMARY KEY," +
            "Result INTEGER," +
            "DSum INTEGER," +
            "DDateTime DATETIME," +
            "Sides INTEGER," +
            "Personnage VARCHAR(50)," +
            "Dices INTEGER," +
            "FOREIGN KEY(Personnage) REFERENCES Personnage(Nom));";

        private const int TargetCharactersDbVersion = 22;

        public async Task RunCharactersDbMigrationsAsync()
        {
            if (_doneMigrations)
            {
                return;
            }
            _doneMigrations = true;
            using var db = CharactersDb.Create();
            if (VersionIsAt(db, TargetCharactersDbVersion))
            {
                return;
            }
            await CreateDiceRollTableAsync(db).ConfigureAwait(false);
            UpdateCharactersDbVersion(db, TargetCharactersDbVersion);
        }

        private static bool VersionIsAt(Database db, short v)
        {
            var version = db.Query<Version?>("SELECT * FROM VERSION;");
            return version?.FirstOrDefault()?.Version1 == v;
        }

        private static void UpdateCharactersDbVersion(Database db, short v)
        {
            db.Execute("UPDATE VERSION SET Version=@0 WHERE Version=@1;", v, v - 1);
        }

        public static Task CreateDiceRollTableAsync(Database db)
        {
            var task = db.ExecuteAsync(DiceRollMigration);
            return task;
        }
    }
}