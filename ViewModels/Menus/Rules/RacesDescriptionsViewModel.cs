using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;

using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Interfaces;

using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Rules
{
    public class RacesDescriptionsViewModel : ViewModelBase, IDocumentViewModel, IDbDependantViewModel
    {
        private bool _isBusy = false;

        private ObservableCollection<RacePersonnage> _races = new ObservableCollection<RacePersonnage>();

        private RacePersonnage? _selectedRace;

        public RacesDescriptionsViewModel()
        {
            Previous = new RelayCommand(PreviousMethod);
            Next = new RelayCommand(NextMethod);
        }

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public RelayCommand Next { get; private set; }

        public RelayCommand Previous { get; private set; }

        public ObservableCollection<RacePersonnage> Races { get => _races; private set { Set(nameof(Races), ref _races, value); } }

        public RacePersonnage? SelectedRace { get => _selectedRace; set { Set(nameof(SelectedRace), ref _selectedRace, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            using var progDb = DB.ProgDb.Create();
            Races = new ObservableCollection<RacePersonnage>(await Task.Run(() => progDb.Query<RacePersonnage>("select race,Description from Race order by race").Where(x => string.IsNullOrWhiteSpace(x.Description) == false && x.Description != x.Race)).ConfigureAwait(false));
            if (Races.Any())
            {
                SelectedRace = Races.FirstOrDefault();
            }
            IsBusy = false;
        }

        private void GoToWithStep(int diff)
        {
            if (SelectedRace is null)
            {
                SelectedRace = Races.FirstOrDefault();
                return;
            }
            var index = Races.IndexOf(SelectedRace);
            if (index != -1)
            {
                var newIdex = index += diff;
                if (newIdex > Races.Count - 1)
                {
                    newIdex = 0;
                }
                else if (newIdex < 0)
                {
                    newIdex = Races.Count - 1;
                }
                SelectedRace = Races.ElementAtOrDefault(newIdex);
            }
        }

        private void NextMethod()
        {
            GoToWithStep(1);
        }

        private void PreviousMethod()
        {
            GoToWithStep(-1);
        }
    }
}