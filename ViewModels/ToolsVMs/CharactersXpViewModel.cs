using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;
using RoleDDNG.DatabaseLayer;
using RoleDDNG.Models.Characters;
using RoleDDNG.Models.Options;
using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Constants;
using RoleDDNG.ViewModels.Extensions;
using RoleDDNG.ViewModels.Interfaces;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
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
            using var charactersDb = new Database(SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath);
            using var progDb = new Database(Consts.ProgramDatabaseFileName);
            var charactersFromDb = await charactersDb.GetQueryDataAsync<Personnage>(DbCharactersQuery).ConfigureAwait(true);
            Characters = new ObservableCollection<Personnage>(charactersFromDb);

            foreach (var character in Characters)
            {
                var races = await progDb.GetQueryDataAsync<Races>($"select AdjNiv from Race where race='{character.Race}'").ConfigureAwait(true);
                var archetypes = await progDb.GetQueryDataAsync<Archetype>($"select AdjNiv from Archetype where archetype='{character.Archetype}'").ConfigureAwait(true);
                var dons = await charactersDb.GetQueryDataAsync<PersonnageDons>($"select dons from PersonnageDons where nom='{character.Nom}'").ConfigureAwait(true);
                foreach (var level in new short?[] { character.Niv1, character.Niv2, character.Niv3, character.Niv4, character.Niv5, character.Niv6, character.Niv7, character.Niv8 })
                {
                    if (level.HasValue)
                    {
                        character.Niveau += level.Value;
                    }
                }
                foreach (var level in new short?[] { character.Niv1, character.Niv2, character.Niv3, character.Niv4, character.Niv5, character.Niv6, character.Niv7, character.Niv8, races.Select(x => x.AdjNiv).FirstOrDefault(), archetypes.Select(x => x.AdjNiv).FirstOrDefault() })
                {
                    if (level.HasValue)
                    {
                        character.NiveauGE += level.Value;
                    }
                }
                if (dons.Select(x => x.Dons).Any(x => x == Consts.LevelOneAdjustmentReduction))
                {
                    character.NiveauGE -= 1;
                }
                if (character.TotalXp.HasValue && character.TotalXp > character.GetXpLevel())
                {
                    character.Xp = character.TotalXp.Value;
                }
                else
                {
                    character.Xp = character.GetXpLevel();
                }
                character.NiveauSuivant = character.GetNextLevelXp();
                if (character.Xp >= character.NiveauSuivant)
                {
                    character.IsBelowNormalXpLevel = true;
                }
            }

            IsBusy = false;
        }
    }
}