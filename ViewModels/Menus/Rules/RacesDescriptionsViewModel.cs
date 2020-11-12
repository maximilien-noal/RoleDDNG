using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Interfaces;

using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Rules
{
    public class RacesDescriptionsViewModel : CollectionSlideShowViewModel<RacePersonnage>, IDocumentViewModel, IDbDependentViewModel
    {
        private bool _isBusy;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            using var progDb = DB.ProgDb.Create();
            using var elementsReader = await progDb.QueryAsync<RacePersonnage>("select race,Description from Race order by race").ConfigureAwait(true);
            var collection = new ObservableCollection<RacePersonnage>();
            while (await elementsReader.ReadAsync().ConfigureAwait(true))
            {
                if (string.IsNullOrWhiteSpace(elementsReader.Poco.Description) == false && elementsReader.Poco.Description != elementsReader.Poco.Race)
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