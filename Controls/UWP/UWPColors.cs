using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

using Windows.UI.ViewManagement;

namespace Hammer.MDI.Control.UWP
{
    public class UWPColors : INotifyPropertyChanged
    {
        private static UWPColors? _instance;

        private static UISettings _uiSettings = new UISettings();

        private SolidColorBrush _accentColor = GetAccentColor();

        private UWPColors()
        {
            _uiSettings.ColorValuesChanged += (s, e) => CopyCurrentAccentColorFromWindows();
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public static UWPColors Instance
        {
            get
            {
                if (_instance is null)
                {
                    _instance = new UWPColors();
                }
                return _instance;
            }
        }

        public SolidColorBrush AccentColor
        {
            get
            {
                return _accentColor;
            }

            set
            {
                if (value != null && _accentColor != value)
                {
                    _accentColor = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public void NotifyPropertyChanged([CallerMemberName] string propertyName = "") => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private static SolidColorBrush GetAccentColor() => new SolidColorBrush(Color.FromArgb(_uiSettings.GetColorValue(UIColorType.Accent).A, _uiSettings.GetColorValue(UIColorType.Accent).R, _uiSettings.GetColorValue(UIColorType.Accent).G, _uiSettings.GetColorValue(UIColorType.Accent).B));

        private void CopyCurrentAccentColorFromWindows() => AccentColor = GetAccentColor();
    }
}