using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.Models.Characters;
using RoleDDNG.ViewModels.DB;
using RoleDDNG.ViewModels.Interfaces;

using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus
{
    public class CharactersListViewModel : CollectionViewModel<Personnage>, IDbDependentViewModel, IDocumentViewModel
    {
        public string Title => "Accéder à un perso";
        private readonly string _sourceDbFile = "";
        private bool _isBusy;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public CharactersListViewModel(string sourceDbFile = "")
        {
            _sourceDbFile = string.IsNullOrWhiteSpace(sourceDbFile) ? SimpleIoc.Default.GetInstance<MainViewModel>().CurrentCharacterDb : sourceDbFile;
        }

        public virtual async Task LoadDbDataAsync()
        {
            if (string.IsNullOrWhiteSpace(_sourceDbFile))
            {
                return;
            }
            await LoadDbDataFromFileAsync(_sourceDbFile).ConfigureAwait(true);
        }

        protected async Task LoadDbDataFromFileAsync(string source, string customQuery = "")
        {
            try
            {
                IsBusy = true;
                Collection = await DatabaseWrapper.GetCollectionFromQueryAsync<Personnage, ObservableCollection<Personnage>>(string.IsNullOrWhiteSpace(customQuery) ? CommonQueries.DbCharactersQuick : customQuery, source).ConfigureAwait(true);
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