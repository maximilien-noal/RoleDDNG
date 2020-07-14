using Hammer.MDI.Control;
using System.Windows.Controls;
using System.Windows.Media;

namespace RoleDDNG.Role.UserControls
{
    public class WindowContent : UserControl
    {
        public WindowContent()
        {
            this.Loaded += WindowContent_Loaded;
        }

        private void WindowContent_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            var imageBrush = (ImageBrush)this.FindResource("WindowIcon");
            if (imageBrush is null)
            {
                return;
            }
            var parent = VisualTreeHelper.GetParent(this);
            while (!(parent is MdiWindow))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }
            if (parent is MdiWindow window)
            {
                window.ChangeMenuButtonIcon(imageBrush);
            }
        }
    }
}