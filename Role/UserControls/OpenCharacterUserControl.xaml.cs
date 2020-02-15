using System.Threading.Tasks;
using System.Windows.Controls;

using AsyncAwaitBestPractices;

using RoleDDNG.ViewModels.ToolsVMs;

namespace RoleDDNG.Role.UserControls
{
    /// <summary>
    /// Interaction logic for OpenCharacterUserControl.xaml
    /// </summary>
    public partial class OpenCharacterUserControl : UserControl
    {
        public OpenCharacterUserControl()
        {
            InitializeComponent();
        }

        private async Task AskForDBFilePathAsync()
        {
            if (DataContext is OpenCharacterViewModel viewModel)
            {
                await viewModel.AskForDatabaseFileCommand.ExecuteAsync().ConfigureAwait(true);
            }
        }

        private void UserControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            AskForDBFilePathAsync().SafeFireAndForget();
        }
    }
}