using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    public class ChangeNomRace : ObservableObject
    {
        private string? _newRace;

        private string? _oldRace;

        public ChangeNomRace()
        {
        }

        [Column]
        public string? NewRace { get => _newRace; set { Set(nameof(NewRace), ref _newRace, value); } }

        [Column]
        public string? OldRace { get => _oldRace; set { Set(nameof(OldRace), ref _oldRace, value); } }
    }
}