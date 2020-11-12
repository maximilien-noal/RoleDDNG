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
            using var progDb = DB.ProgDb.Create();
            using var elementsReader = await progDb.QueryAsync<Don>("select * from Dons order by Nom").ConfigureAwait(true);
            var collection = new ObservableCollection<Don>();
            while (await elementsReader.ReadAsync().ConfigureAwait(true))
            {
                if (string.IsNullOrWhiteSpace(elementsReader.Poco.Nom) == false)
                {
                    collection.Add(elementsReader.Poco);
                }
            }
            SetCollection(collection);
            if (Collection.Any())
            {
                SelectedItem = Collection.First();
            }
            IsBusy = false;
        }
    }
}