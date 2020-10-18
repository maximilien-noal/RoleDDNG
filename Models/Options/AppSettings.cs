using RoleDDNG.Models.Structs;

namespace RoleDDNG.Models.Options
{
    public class AppSettings : ObservableObject
    {
        private string _lastCharacterDBPath = "";

        private WindowPlacement _mainWindowPlacement;

        public string LastCharacterDBPath { get => _lastCharacterDBPath; set { Set(nameof(LastCharacterDBPath), ref _lastCharacterDBPath, value); } }

        public WindowPlacement MainWindowPlacement { get => _mainWindowPlacement; set { Set(nameof(MainWindowPlacement), ref _mainWindowPlacement, value); } }
    }
}