using GalaSoft.MvvmLight;
using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Interfaces;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.DescriptionsVMs
{
    public class RacesDescriptionsViewModel : ViewModelBase, IDocumentViewModel, IDbDependantViewModel
    {
        private bool _isBusy = false;

        private ObservableCollection<RacePersonnage> _races = new ObservableCollection<RacePersonnage>();

        private RacePersonnage? _selectedRace;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public ObservableCollection<RacePersonnage> Races { get => _races; private set { Set(nameof(Races), ref _races, value); } }

        public RacePersonnage? SelectedRace { get => _selectedRace; set { Set(nameof(SelectedRace), ref _selectedRace, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            using var progDb = DB.ProgDbInstanceCreator.Create();
            var races = new List<RacePersonnage>();
            await progDb.QueryAsync<RacePersonnage>(x => races.Add(x), "select race,Description from Race order by race").ConfigureAwait(true);
            Races = new ObservableCollection<RacePersonnage>(races.Where(x => string.IsNullOrWhiteSpace(x.Description) == false));
            if (Races.Any())
            {
                SelectedRace = Races.FirstOrDefault();
            }
            IsBusy = false;
        }
    }
}