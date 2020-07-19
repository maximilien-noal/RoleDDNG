using GalaSoft.MvvmLight.Ioc;
using Hammer.MDI.Control;
using RoleDDNG.ViewModels;
using RoleDDNG.ViewModels.Interfaces;
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

#pragma warning disable VSTHRD100 // Avoid async void methods (this is an event)

        private async void WindowContent_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            if (DataContext is ICharactersDbDependentViewModel viewModel)
            {
                await viewModel.LoadCharactersDbDataAsync().ConfigureAwait(true);
            }

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

#pragma warning restore VSTHRD100 // Avoid async void methods (this is an event)
    }
}