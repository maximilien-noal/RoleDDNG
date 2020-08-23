using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Interfaces;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Rules
{
    public class DonsDescriptionsViewModel : CollectionViewModel<Don>, IDocumentViewModel, IDbDependantViewModel
    {
        private bool _isBusy;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            using var progDb = DB.ProgDb.Create();
            SetCollection(new ObservableCollection<Don>(await Task.Run(() => progDb.Query<Don>("select * from Dons order by Nom").Where(x => string.IsNullOrWhiteSpace(x.Nom) == false)).ConfigureAwait(false)));
            if (Collection.Any())
            {
                SelectedItem = Collection.FirstOrDefault();
            }
            IsBusy = false;
        }
    }
}