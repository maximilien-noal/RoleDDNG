using GalaSoft.MvvmLight;

namespace RoleDDNg.Models
{
    public class AppSettings : ObservableObject
    {
        private string _mainWindowPlacement = "";
        public string MainWindowPlacement { get => _mainWindowPlacement; set { Set(nameof(MainWindowPlacement), ref _mainWindowPlacement, value); } }
    }
}