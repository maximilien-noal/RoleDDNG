using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Interfaces;

using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Rules
{
    public class RacesDescriptionsViewModel : CollectionSlideShowViewModel<RacePersonnage>, IDocumentViewModel, IDbDependentViewModel
    {
        public string Title => "Desc. des races";
        private bool _isBusy;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            var collection = await DB.DatabaseWrapper.GetCollectionFromQueryAsync<RacePersonnage, List<RacePersonnage>>("select race,Description from Race order by race", DB.DatabaseWrapper.GetFullProgDbPath()).ConfigureAwait(true);
            SetCollection(new ObservableCollection<RacePersonnage>(collection.Where(x => string.IsNullOrWhiteSpace(x.Description) == false && x.Description != x.Race)));
            if (Collection.Any())
            {
                SelectedItem = Collection.First();
            }
            IsBusy = false;
        }
    }
}