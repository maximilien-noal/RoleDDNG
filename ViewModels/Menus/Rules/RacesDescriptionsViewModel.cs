﻿using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Interfaces;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Rules
{
    public class RacesDescriptionsViewModel : CollectionViewModel<RacePersonnage>, IDocumentViewModel, IDbDependantViewModel
    {
        private bool _isBusy;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            using var progDb = DB.ProgDb.Create();
            SetCollection(new ObservableCollection<RacePersonnage>(await Task.Run(() => progDb.Query<RacePersonnage>("select race,Description from Race order by race").Where(x => string.IsNullOrWhiteSpace(x.Description) == false && x.Description != x.Race)).ConfigureAwait(false)));
            if (Collection.Any())
            {
                SelectedItem = Collection.FirstOrDefault();
            }
            IsBusy = false;
        }
    }
}