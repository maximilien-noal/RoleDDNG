using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;
using RoleDDNG.DatabaseLayer;
using RoleDDNG.Models.Characters;
using RoleDDNG.Models.Options;
using RoleDDNG.ViewModels.Interfaces;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.ToolsVMs
{
    public class CharactersXpViewModel : ViewModelBase, IDocumentViewModel, ICharactersDbDependentViewModel
    {
        private const string DbCharactersQuery = "select nom,image,race,niv_1,niv_2,niv_3,niv_4,niv_5,niv_6,niv_7,niv_8,malusxp,archetype,totalxp,classe_1,classe_2,classe_3,classe_4,classe_5,classe_6,classe_7,classe_8 from personnage where exclu=false order by nom";

        private ObservableCollection<Personnage> _characters = new ObservableCollection<Personnage>();

        private bool _isBusy = false;

        private Personnage? _selectedCharacter;

        private int _xpCalculated = 0;

        private double _xpPercentage = 0;

        private int _xpToAdd = 0;

        public CharactersXpViewModel()
        {
            Cancel = new RelayCommand(() => SimpleIoc.Default.GetInstance<MainViewModel>().RemoveMdiWindow<CharactersXpViewModel>());
        }

        public RelayCommand Cancel { get; private set; }

        public ObservableCollection<Personnage> Characters { get => _characters; private set { Set(nameof(Characters), ref _characters, value); } }

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public Personnage? SelectedCharacter { get => _selectedCharacter; set { Set(nameof(SelectedCharacter), ref _selectedCharacter, value); } }

        public int XpCalculated { get => _xpCalculated; set { Set(nameof(XpCalculated), ref _xpCalculated, value); } }

        public double XpPercentage { get => _xpPercentage; set { Set(nameof(XpPercentage), ref _xpPercentage, value); } }

        public int XpToAdd { get => _xpToAdd; set { Set(nameof(XpToAdd), ref _xpToAdd, value); } }

        public async Task LoadCharactersDbDataAsync()
        {
            IsBusy = true;
            var db = new Database(SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath);
            var charactersFromDb = await db.GetQueryDataAsync<Personnage>(DbCharactersQuery).ConfigureAwait(true);
            Characters = new ObservableCollection<Personnage>(charactersFromDb);
            IsBusy = false;
        }
    }
}