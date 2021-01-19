using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Interfaces;

using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Rules
{
    public class DonsDescriptionsViewModel : CollectionSlideShowViewModel<Don>, IDocumentViewModel, IDbDependentViewModel
    {
        private bool _isBusy;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            var collection = await DB.DatabaseWrapper.GetCollectionFromQueryAsync<Don, ObservableCollection<Don>>("select * from Dons order by Nom", DB.DatabaseWrapper.GetFullProgDbPath()).ConfigureAwait(true);
            SetCollection(collection);
            if (Collection.Any())
            {
                SelectedItem = Collection.First();
            }
            IsBusy = false;
        }
    }
}