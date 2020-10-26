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

        private void WindowContent_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            var vm = this.DataContext as CharacterImportViewModel;
            if (vm != null)
            {
                vm.SetImportNames();
            }
        }
    }
}