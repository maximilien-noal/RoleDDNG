using GalaSoft.MvvmLight;

using RoleDDNG.Models.Structs;

namespace RoleDDNG.Models.Options
{
    public class AppSettings : ObservableObject
    {
        private string _bdd = default;

        private WindowPlacement _mainWindowPlacement = default;

        public string Bdd { get => _bdd; set { Set(nameof(_bdd), ref _bdd, value); } }

        public WindowPlacement MainWindowPlacement { get => _mainWindowPlacement; set { Set(nameof(MainWindowPlacement), ref _mainWindowPlacement, value); } }
    }
}