using GalaSoft.MvvmLight;

using RoleDDNG.Models.Characters;
using RoleDDNG.ViewModels.DB;
using RoleDDNG.ViewModels.Interfaces;

using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Tools
{
    public class OpenCharacterViewModel : ViewModelBase, IDocumentViewModel, IDbDependentViewModel
    {
        private ObservableCollection<Personnage> _characters = new ObservableCollection<Personnage>();

        private bool _isBusy;

        public ObservableCollection<Personnage> Characters { get => _characters; private set { Set(nameof(Characters), ref _characters, value); } }

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            using var charactersDb = DB.CharactersDb.Create();
            Characters = new ObservableCollection<Personnage>(await Task.Run(() => charactersDb.Query<Personnage>(CommonQueries.DbCharactersQuery)).ConfigureAwait(true));
            IsBusy = false;
        }
    }
}