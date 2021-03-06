﻿using AsyncAwaitBestPractices.MVVM;

using RoleDDNG.Models.Characters;
using RoleDDNG.Models.Roles;
using RoleDDNG.ViewModels.Constants;
using RoleDDNG.ViewModels.Interfaces;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Tools
{
    public class CharactersXpViewModel : ViewModelWithCloseAction<CharactersXpViewModel>, IDocumentViewModel, IDbDependentViewModel
    {
        private const string DbCharactersQuery = "select nom,image,race,niv_1,niv_2,niv_3,niv_4,niv_5,niv_6,niv_7,niv_8,malusxp,archetype,totalxp,classe_1,classe_2,classe_3,classe_4,classe_5,classe_6,classe_7,classe_8 from personnage where exclu=false order by nom";

        private ObservableCollection<Personnage> _characters = new ObservableCollection<Personnage>();

        private ObservableCollection<Personnage> _charactersLog = new ObservableCollection<Personnage>();

        private int _fp;

        private bool _isBusy;

        private Personnage? _selectedCharacter;

        private double _xpCalculated;

        private double _xpPercentage;

        public CharactersXpViewModel()
        {
            Add = new AsyncCommand(AddMethodAsync);
            Save = new AsyncCommand(SaveMethodAsync);
        }

        public AsyncCommand Add { get; private set; }

        public ObservableCollection<Personnage> Characters { get => _characters; private set { Set(nameof(Characters), ref _characters, value); } }

        public ObservableCollection<Personnage> CharactersLog { get => _charactersLog; private set { Set(nameof(CharactersLog), ref _charactersLog, value); } }

        public int FP { get => _fp; set { Set(nameof(FP), ref _fp, value); } }

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public AsyncCommand Save { get; private set; }

        public Personnage? SelectedCharacter { get => _selectedCharacter; set { Set(nameof(SelectedCharacter), ref _selectedCharacter, value); } }

        public double XpCalculated { get => _xpCalculated; set { Set(nameof(XpCalculated), ref _xpCalculated, value); } }

        public double XpPercentage { get => _xpPercentage; set { Set(nameof(XpPercentage), ref _xpPercentage, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;

            var charactersRaces = new List<RacePersonnage>();
            var charactersArchetypes = new List<Archetype>();
            var charactersGifts = new List<PersonnageDons>();
            var characters = new List<Personnage>();

            var task1 = Task.Run(() =>
            {
                using var progDb = DB.DatabaseWrapper.CreateProgDb();
                charactersRaces.AddRange(progDb.Query<RacePersonnage>("select AdjNiv,race from Race"));
            });
            var task2 = Task.Run(() =>
            {
                using var progDb = DB.DatabaseWrapper.CreateProgDb();
                charactersArchetypes.AddRange(progDb.Query<Archetype>("select AdjNiv,archetype from Archetype"));
            });
            var task3 = Task.Run(() =>
            {
                using var charactersDb = DB.DatabaseWrapper.CreateCharactersDb();
                charactersGifts.AddRange(charactersDb.Query<PersonnageDons>("select nom,dons from PersonnageDons"));
            });
            var task4 = Task.Run(() =>
            {
                using var charactersDb = DB.DatabaseWrapper.CreateCharactersDb();
                characters.AddRange(charactersDb.Query<Personnage>(DbCharactersQuery));
            });

            await Task.WhenAll(task1, task2, task3, task4).ConfigureAwait(true);

            Characters = new ObservableCollection<Personnage>(characters);

            foreach (var character in Characters)
            {
                var race = charactersRaces.Where(x => x.Race == character.Race).FirstOrDefault();
                var archetype = charactersArchetypes.Where(x => x.NomArchetype == character.Archetype).FirstOrDefault();
                var don = charactersGifts.Where(x => x.Nom == character.Nom).FirstOrDefault();
                foreach (var level in new short?[] { character.Niv1, character.Niv2, character.Niv3, character.Niv4, character.Niv5, character.Niv6, character.Niv7, character.Niv8 })
                {
                    if (level.HasValue)
                    {
                        character.Niveau += level.Value;
                    }
                }
                foreach (var level in new short?[] { character.Niv1, character.Niv2, character.Niv3, character.Niv4, character.Niv5, character.Niv6, character.Niv7, character.Niv8, (short?)race?.AdjNiv, archetype?.AdjNiv })
                {
                    if (level.HasValue)
                    {
                        character.NiveauGE += level.Value;
                    }
                }
                if (don?.Dons == Consts.LevelOneAdjustmentReduction)
                {
                    character.NiveauGE -= 1;
                }
                if (character.TotalXp.HasValue && character.TotalXp > character.GetXpPointsForLevel())
                {
                    character.Xp = character.TotalXp.Value;
                }
                else
                {
                    character.Xp = character.GetXpPointsForLevel();
                }
                character.NiveauSuivant = character.GetXpPointsForNextLevel();
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
            FP = Math.Max(Math.Min(FP, 125), 0);
            var experience = new List<Experience>(await Task.Run(() => DB.DatabaseWrapper.CreateCharactersDb(DB.DatabaseWrapper.GetFullProgDbPath()).Query<Experience>($"select fp{FP} from experience where niveau=@0", SelectedCharacter.NiveauGE)).ConfigureAwait(true));
            if (experience.Any() == false)
            {
                return;
            }
            CapXpPercentage();
            var fpValue = (double?)experience.FirstOrDefault()?.GetType().GetProperty($"FP{FP}")?.GetValue(experience.FirstOrDefault(), null);
            if (fpValue is null || fpValue.HasValue == false)
            {
                return;
            }
            SelectedCharacter.GainXp = SelectedCharacter.Part * fpValue.Value * (1 - XpPercentage / 100);
        }

        private void CapXpPercentage()
        {
            XpPercentage = Math.Max(Math.Min(XpPercentage, 100), 0);
        }

        private async Task SaveMethodAsync()
        {
            if (SelectedCharacter is null)
            {
                return;
            }
            SelectedCharacter.Xp += SelectedCharacter.GainXp;
            CapXpPercentage();
            XpCalculated += XpCalculated + SelectedCharacter.GainXp * XpPercentage / 100 - XpPercentage;
            SelectedCharacter.GainXp = 0;
            SelectedCharacter.TotalXp = (long?)Math.Min(SelectedCharacter.Xp, long.MaxValue);
            CharactersLog.Add(SelectedCharacter);
            await Task.Run(() =>
            {
                using var charactersDb = DB.DatabaseWrapper.CreateCharactersDb();
                return charactersDb.Update(SelectedCharacter, SelectedCharacter.Nom, columns: new string[] { "TotalXP" });
            }).ConfigureAwait(true);
        }
    }
}