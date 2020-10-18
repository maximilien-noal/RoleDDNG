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
            "Id AUTOINCREMENT PRIMARY KEY," +
            "Results MEMO," +
            "DSum INTEGER," +
            "DDateTime DATETIME," +
            "Sides INTEGER," +
            "Dices INTEGER," +
            "Personnage VARCHAR(50));";

        private const int TargetFinalDbVersion = 22;

        public Task RunCharactersDbMigrationsAsync()
        {
            if (_doneMigrations)
            {
                return Task.Delay(0);
            }
            _doneMigrations = true;
            return CreateDiceRollTableAsync();
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

        public static Task CreateDiceRollTableAsync()
        {
            var task = Task.Run(() =>
            {
                using var db = CharactersDb.Create();
                if (VersionIsAt(db, 22))
                {
                    return;
                }
                db.Execute(DiceRollMigration);
                UpdateCharactersDbVersion(db, 22);
            });

            return task;
        }

        internal static bool NeedsToRun()
        {
            using var db = CharactersDb.Create();
            return VersionIsAt(db, TargetFinalDbVersion) == false;
        }
    }
}