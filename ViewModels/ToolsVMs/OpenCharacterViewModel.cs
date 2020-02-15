using System.IO;
using System.Threading.Tasks;

using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.Interfaces.Dialogs;
using RoleDDNG.ViewModels.Interfaces;

namespace RoleDDNG.ViewModels.ToolsVMs
{
    public class OpenCharacterViewModel : ViewModelBase, IContent
    {
        private string _databasePath;

        public OpenCharacterViewModel()
        {
            AskForDatabaseFileCommand = new AsyncCommand(AskForDatabaseFileAsync);
        }

        public AsyncCommand AskForDatabaseFileCommand { get; }

        public string Title => "Accéder à un personnage";

        private async Task AskForDatabaseFileAsync()
        {
            var fileDialog = SimpleIoc.Default.GetInstance<IFileDialog>();
            var dbFile = await fileDialog.TryOpenUserChosenFileAsync("Ouvrir une base de données de personnages...", "mdb").ConfigureAwait(true);
            if (string.IsNullOrWhiteSpace(dbFile) == false && File.Exists(dbFile))
            {
                _databasePath = dbFile;
            }
            else
            {
                SimpleIoc.Default.GetInstance<MainViewModel>().RemoveMdiWindow<OpenCharacterViewModel>();
            }
            SimpleIoc.Default.GetInstance<MainViewModel>().SetCharacterDatabasePath.Execute(_databasePath);
        }
    }
}