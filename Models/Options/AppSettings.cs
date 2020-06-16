using GalaSoft.MvvmLight;

using RoleDDNG.Models.Structs;

namespace RoleDDNG.Models.Options
{
    public class AppSettings : ObservableObject
    {
        private string _lastCharacterDBPath = "";

        private WindowPlacement _mainWindowPlacement = default;

        public string LastCharacterDBPath { get => _lastCharacterDBPath; set { Set(nameof(_lastCharacterDBPath), ref _lastCharacterDBPath, value); } }

        public WindowPlacement MainWindowPlacement { get => _mainWindowPlacement; set { Set(nameof(MainWindowPlacement), ref _mainWindowPlacement, value); } }
    }
}