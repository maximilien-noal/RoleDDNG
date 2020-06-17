using Hammer.MDI.Control;
using System;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace RoleDDNG.Role.UserControls.Traits
{
    internal interface IUpdateMdiWindowIconTrait
    {
        public void ChangeMdiWindowIcon(UserControl userControl)
        {
            var imageBrush = (ImageBrush)userControl.FindResource("WindowIcon");
            if (imageBrush is null)
            {
                return;
            }
            var parent = VisualTreeHelper.GetParent(userControl);
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