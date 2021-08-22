using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Interfaces;

using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Rules
{
    public class SpellsDescriptionsViewModel : CollectionSlideShowViewModel<Spell>, IDocumentViewModel, IDbDependentViewModel
    {
        public string Title => "Desc. des sorts";
        private bool _isBusy;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            var collection = await DB.DatabaseWrapper.GetCollectionFromQueryAsync<Spell, ObservableCollection<Spell>>("select * from Sort order by Nom", DB.DatabaseWrapper.GetFullProgDbPath()).ConfigureAwait(true);
            SetCollection(collection);
            if (Collection.Any())
            {
                SelectedItem = Collection.First();
            }
            IsBusy = false;
        }
    }
}