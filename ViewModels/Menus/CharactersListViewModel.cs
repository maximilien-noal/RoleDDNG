using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.Models.Characters;
using RoleDDNG.ViewModels.DB;
using RoleDDNG.ViewModels.Interfaces;

using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus
{
    public class CharactersListViewModel : CollectionViewModel<Personnage>, IDbDependentViewModel, IDocumentViewModel
    {
        private string _sourceDbFile = "";
        private bool _isBusy;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public CharactersListViewModel(string sourceDbFile = "")
        {
            _sourceDbFile = string.IsNullOrWhiteSpace(sourceDbFile) ? SimpleIoc.Default.GetInstance<MainViewModel>().CurrentCharacterDb : sourceDbFile;
        }
        public async Task LoadDbDataAsync()
        {
            try
            {
                IsBusy = true;
                using var charactersDb = CharactersDb.Create(_sourceDbFile);
                using var elementsReader = await charactersDb.QueryAsync<Personnage>(CommonQueries.DbCharactersAll).ConfigureAwait(true);
                while (await elementsReader.ReadAsync().ConfigureAwait(true))
                {
                    Collection.Add(elementsReader.Poco);
                }
                IsBusy = false;
            }
            catch
            {
                Cancel.Execute(this);
                throw;
            }
        }
    }
}
