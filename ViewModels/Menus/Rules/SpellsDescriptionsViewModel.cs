using GalaSoft.MvvmLight;
using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Interfaces;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Rules
{
    public class SpellsDescriptionsViewModel : ViewModelBase, IDocumentViewModel, IDbDependantViewModel
    {
        private bool _isBusy = false;

        private Sort? _selectedSpell;

        private ObservableCollection<Sort> _spells = new ObservableCollection<Sort>();

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public Sort? SelectedSpell { get => _selectedSpell; set { Set(nameof(SelectedSpell), ref _selectedSpell, value); } }

        public ObservableCollection<Sort> Spells { get => _spells; private set { Set(nameof(Spells), ref _spells, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            using var progDb = DB.ProgDb.Create();
            Spells = new ObservableCollection<Sort>(await Task.Run(() => progDb.Query<Sort>("select Nom,Explication from Sort order by Nom").Where(x => string.IsNullOrWhiteSpace(x.Explication) == false)).ConfigureAwait(false));
            if (Spells.Any())
            {
                SelectedSpell = Spells.FirstOrDefault();
            }
            IsBusy = false;
        }
    }
}