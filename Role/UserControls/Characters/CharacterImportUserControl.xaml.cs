using RoleDDNG.ViewModels.Menus.Characters;

namespace RoleDDNG.Role.UserControls
{
    /// <summary>
    /// Interaction logic for CharacterImportUserControl.xaml
    /// </summary>
    public partial class CharacterImportUserControl : WindowContent
    {
        public CharacterImportUserControl()
        {
            InitializeComponent();
        }

#pragma warning disable VSTHRD100 // Avoid async void methods (this is an event)

        private async void WindowContent_Loaded(object sender, System.Windows.RoutedEventArgs e)
#pragma warning restore VSTHRD100 // Avoid async void methods (this is an event)
        {
            if (this.DataContext is CharacterImportViewModel vm)
            {
                await vm.SetImportNamesAsync().ConfigureAwait(true);
            }
        }
    }
}