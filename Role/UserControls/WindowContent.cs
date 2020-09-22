using Hammer.MDI.Control;
using RoleDDNG.ViewModels.Interfaces;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace RoleDDNG.Role.UserControls
{
    public class WindowContent : UserControl
    {
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(nameof(Title), typeof(string), typeof(WindowContent), new PropertyMetadata(""));

        public WindowContent()
        {
            this.Loaded += WindowContent_Loaded;
        }

        public string Title
        {
            get { return (string)GetValue(TitleProperty); }
            set { SetValue(TitleProperty, value); }
        }

#pragma warning disable VSTHRD100 // Avoid async void methods (this is an event)

        private async void WindowContent_Loaded(object sender, RoutedEventArgs e)
        {
            if (DesignerProperties.GetIsInDesignMode(this))
            {
                return;
            }

            if (DataContext is IDbDependantViewModel viewModel)
            {
                await viewModel.LoadDbDataAsync().ConfigureAwait(true);
            }

            var imageBrush = this.FindResource("WindowIcon");
            if (imageBrush is null)
            {
                return;
            }
            var parent = VisualTreeHelper.GetParent(this);
            while (!(parent is MdiWindow))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }
            if (parent is MdiWindow window && imageBrush is ImageBrush brush)
            {
                window.SetCurrentValue(MdiWindow.TitleProperty, GetValue(TitleProperty));
                window.ChangeMenuButtonIcon(brush);
            }
        }

#pragma warning restore VSTHRD100 // Avoid async void methods (this is an event)
    }
}