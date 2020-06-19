using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Ioc;
using RoleDDNG.DatabaseLayer;
using RoleDDNG.Interfaces.Dialogs;
using RoleDDNG.Models.Characters;
using RoleDDNG.ViewModels.Interfaces;

namespace RoleDDNG.ViewModels.ToolsVMs
{
    public class OpenCharacterViewModel : ViewModelBase, IDocumentViewModel
    {
        private const string DbCharactersQuery = "select nom,image,race,niv_1,niv_2,niv_3,niv_4,niv_5,niv_6,niv_7,niv_8,classe_1,classe_2,classe_3,classe_4,classe_5,classe_6,classe_7,classe_8 from personnage where exclu=false order by nom";

        private bool _isBusy = false;

        public OpenCharacterViewModel()
        {
            AskForDatabaseFileCommand = new AsyncCommand(AskForDatabaseFileAsync);
        }

        public AsyncCommand AskForDatabaseFileCommand { get; }

        public ObservableCollection<Personnage> Characters { get; private set; } = new ObservableCollection<Personnage>();

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public string Title => "Accéder à un personnage";

        public async Task GetCharactersFromDbAsync(string dbFile)
        {
            IsBusy = true;
            var db = new DbAccessor(dbFile);
            var charactersFromDb = await db.GetQueryDataAsync<Personnage>(DbCharactersQuery).ConfigureAwait(true);

            if (Characters.Any())
            {
                Characters.Clear();
            }

            charactersFromDb.ToList().ForEach(x => Characters.Add(x));
            IsBusy = false;
        }

        private async Task AskForDatabaseFileAsync()
        {
            IsBusy = true;
            var fileDialog = SimpleIoc.Default.GetInstance<IFileDialog>();
            var dbFile = await fileDialog.OpenFileDialogAsync("Ouvrir une base de données de personnages...", "mdb").ConfigureAwait(true);
            if (!string.IsNullOrWhiteSpace(dbFile) && File.Exists(dbFile))
            {
                SimpleIoc.Default.GetInstance<MainViewModel>().SetCharacterDatabasePath.Execute(dbFile);
            }
            else
            {
                SimpleIoc.Default.GetInstance<MainViewModel>().RemoveMdiWindow<OpenCharacterViewModel>();
            }
            await GetCharactersFromDbAsync(dbFile).ConfigureAwait(true);
            IsBusy = false;
        }
    }
}