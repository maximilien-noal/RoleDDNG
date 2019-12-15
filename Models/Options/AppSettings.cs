using GalaSoft.MvvmLight;

using RoleDDNG.Models.Structs;

using System;

namespace RoleDDNG.Models.Options
{
    public class AppSettings : ObservableObject
    {
        private WindowPlacement _mainWindowPlacement = default;

        public WindowPlacement MainWindowPlacement { get => _mainWindowPlacement; set { Set(nameof(MainWindowPlacement), ref _mainWindowPlacement, value); } }

        private Uri _bdd = default;

        public Uri Bdd { get => _bdd; set { Set(nameof(_bdd), ref _bdd, value); } }
    }
}