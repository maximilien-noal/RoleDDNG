using RoleDDNG.ViewModels.Interfaces;

using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;

namespace RoleDDNG.Role.UserControls
{
    public class WindowContent : UserControl
    {
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(nameof(Title), typeof(string), typeof(WindowContent), new PropertyMetadata(""));

        public WindowContent() => this.Loaded += WindowContent_Loaded;

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

            if (DataContext is IDbDependentViewModel viewModel)
            {
                await viewModel.LoadDbDataAsync().ConfigureAwait(true);
            }
        }

#pragma warning restore VSTHRD100 // Avoid async void methods (this is an event)
    }
}