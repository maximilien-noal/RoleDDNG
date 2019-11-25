using GalaSoft.MvvmLight;
using System;

namespace RoleDDNG.Models.Options
{
    /// <summary>
    /// UserSecrets in %WINDIR%/ROLE.INI
    /// </summary>
    public class AppSettings : ObservableObject
    {
        private string _mainWindowX = "";

        public string MainWindowX { get => _mainWindowX; set { Set(nameof(MainWindowX), ref _mainWindowX, value); } }

        private string _mainWindowY = "";

        public string MainWindowY { get => _mainWindowY; set { Set(nameof(MainWindowY), ref _mainWindowY, value); } }

        private Uri _bdd = default;

        public Uri Bdd { get => _bdd; set { Set(nameof(_bdd), ref _bdd, value); } }
    }
}