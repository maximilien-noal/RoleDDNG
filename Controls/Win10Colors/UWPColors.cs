using System.ComponentModel;
using System.Windows.Media;

using Windows.UI.ViewManagement;

namespace Hammer.MDI.Control.Win10Colors
{
    public static class UWPColors
    {
        public static SolidColorBrush AccentColor
        {
            get
            {
                var uiSettings = new UISettings();
                var accentColor = uiSettings.GetColorValue(UIColorType.Accent);
                var colorBrush = new SolidColorBrush(Color.FromArgb(accentColor.A, accentColor.R, accentColor.G, accentColor.B));
                return colorBrush;
            }
        }
    }
}