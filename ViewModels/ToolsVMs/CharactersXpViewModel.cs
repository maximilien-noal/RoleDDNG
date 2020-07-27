using AsyncAwaitBestPractices.MVVM;
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
using System;
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

        private ObservableCollection<Personnage> _charactersLog = new ObservableCollection<Personnage>();

        private int _fp = 0;

        private bool _isBusy = false;

        private Personnage? _selectedCharacter;

        private int _xpCalculated = 0;

        private double _xpPercentage = 0;

        public CharactersXpViewModel()
        {
            Add = new AsyncCommand(AddMethodAsync);
            Save = new AsyncCommand(SaveMethodAsync);
            Cancel = new RelayCommand(() => SimpleIoc.Default.GetInstance<MainViewModel>().RemoveMdiWindow<CharactersXpViewModel>());
        }

        public AsyncCommand Add { get; private set; }

        public RelayCommand Cancel { get; private set; }

        public ObservableCollection<Personnage> Characters { get => _characters; private set { Set(nameof(Characters), ref _characters, value); } }

        public ObservableCollection<Personnage> CharactersLog { get => _charactersLog; private set { Set(nameof(CharactersLog), ref _charactersLog, value); } }

        public int FP { get => _fp; set { Set(nameof(FP), ref _fp, value); } }

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public AsyncCommand Save { get; private set; }

        public Personnage? SelectedCharacter { get => _selectedCharacter; set { Set(nameof(SelectedCharacter), ref _selectedCharacter, value); } }

        public int XpCalculated { get => _xpCalculated; set { Set(nameof(XpCalculated), ref _xpCalculated, value); } }

        public double XpPercentage { get => _xpPercentage; set { Set(nameof(XpPercentage), ref _xpPercentage, value); } }

        public async Task LoadCharactersDbDataAsync()
        {
            IsBusy = true;

            using var charactersDb = new AccessDb(SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath).GetDatabase();
            using var progDb = new AccessDb(Consts.ProgramDatabaseFileName).GetDatabase();

            var charactersRaces = new List<RacePersonnage>();

            await progDb.QueryAsync<RacePersonnage>(x => charactersRaces.Add(x), "select race,AdjNiv from Race").ConfigureAwait(true);
            var charactersArchetypes = new List<Archetype>();
            await progDb.QueryAsync<Archetype>(x => charactersArchetypes.Add(x), "select archetype,AdjNiv from Archetype").ConfigureAwait(true);
            var charactersGifts = new List<PersonnageDons>();
            await charactersDb.QueryAsync<PersonnageDons>(x => charactersGifts.Add(x), "select nom,dons from PersonnageDons").ConfigureAwait(true);
            if (Characters.Any())
            {
                Characters.Clear();
            }
            await charactersDb.QueryAsync<Personnage>(p => Characters.Add(p), DbCharactersQuery).ConfigureAwait(true);

            foreach (var character in Characters)
            {
                var races = charactersRaces.Where(x => x.Race == character.Race);
                var archetypes = charactersArchetypes.Where(x => x.NomArchetype == character.Archetype);
                var dons = charactersGifts.Where(x => x.Nom == character.Nom);
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

        private async Task AddMethodAsync()
        {
            if (SelectedCharacter is null)
            {
                return;
            }
            using var progDb = new AccessDb(Consts.ProgramDatabaseFileName).GetDatabase();
            var experience = new List<Experience>();
            var fp = Math.Max(Math.Min(FP, 125), 0);
            await progDb.QueryAsync<Experience>((x) => experience.Add(x), $"select fp{fp} from experience where niveau={SelectedCharacter.NiveauGE}").ConfigureAwait(true);
            if (experience.Any() == false)
            {
                return;
            }
            var percentage = Math.Max(Math.Min(XpPercentage, 100), 0);
            var fpValue = (int?)experience.FirstOrDefault()?.GetType().GetProperty($"FP{fp}")?.GetValue(experience.FirstOrDefault(), null);
            if (fpValue is null || fpValue.HasValue == false)
            {
                return;
            }
            SelectedCharacter.GainXp = (long)(SelectedCharacter.Part * fpValue.Value * (1 - percentage / 100));
        }

        private async Task SaveMethodAsync()
        {
            if (SelectedCharacter is null)
            {
                return;
            }
            using var charactersDb = new AccessDb(SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath).GetDatabase();
            await charactersDb.UpdateAsync(SelectedCharacter).ConfigureAwait(true);
        }
    }
}