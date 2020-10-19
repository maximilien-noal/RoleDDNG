using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Interfaces;

using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Rules
{
    public class SpellsDescriptionsViewModel : CollectionViewModel<Spell>, IDocumentViewModel, IDbDependentViewModel
    {
        private bool _isBusy;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            using var progDb = DB.ProgDb.Create();
            SetCollection(new ObservableCollection<Spell>(await Task.Run(() => progDb.Query<Spell>("select * from Sort order by Nom").Where(x => string.IsNullOrWhiteSpace(x.Explication) == false)).ConfigureAwait(true)));
            if (Collection.Any())
            {
                SelectedItem = Collection.FirstOrDefault();
            }
            IsBusy = false;
        }
    }
}