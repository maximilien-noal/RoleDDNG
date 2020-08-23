using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class ChangeNomRace : ObservableObject
    {
        private string? _newRace;

        private string? _oldRace;

        public ChangeNomRace()
        {
        }

        public string? NewRace { get => _newRace; set { Set(nameof(NewRace), ref _newRace, value); } }

        public string? OldRace { get => _oldRace; set { Set(nameof(OldRace), ref _oldRace, value); } }
    }
}