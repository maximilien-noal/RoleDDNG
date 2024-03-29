﻿using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight;

using RandN;
using RandN.Compat;

using RoleDDNG.Models.Characters;
using RoleDDNG.ViewModels.Interfaces;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Tools
{
    public class DiceRollViewModel : ViewModelBase, IDocumentViewModel, IDbDependentViewModel
    {
        public string Title => "Lanceur de dés";

        private const string DbDiceRollsQuery = "SELECT TOP 300 Results,DSum,DDateTime,Sides,Dices FROM DiceRoll WHERE Personnage=\"\" ORDER BY DDateTime DESC;";

        private int _numberOfDices = 1;

        private int _numberOfSides = 6;

        private ObservableCollection<int> _results = new ObservableCollection<int>();

        private ObservableCollection<DiceRoll> _history = new ObservableCollection<DiceRoll>();

        private int _sum;

        public DiceRollViewModel()
        {
            DiceTypes.Add(4);
            DiceTypes.Add(6);
            DiceTypes.Add(8);
            DiceTypes.Add(10);
            DiceTypes.Add(12);
            DiceTypes.Add(20);
            DiceTypes.Add(100);
            Roll = new AsyncCommand(RollAsync);
        }

        public List<int> DiceTypes { get; private set; } = new List<int>();

        private bool _isBusy;

        public bool IsBusy { get => _isBusy; private set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public int NumberOfDices { get => _numberOfDices; set { Set(nameof(NumberOfDices), ref _numberOfDices, value); } }

        public int NumberofSides { get => _numberOfSides; set { Set(nameof(NumberofSides), ref _numberOfSides, value); } }

        public ObservableCollection<int> Results { get => _results; private set { Set(nameof(Results), ref _results, value); } }

        public ObservableCollection<DiceRoll> History { get => _history; private set { Set(nameof(History), ref _history, value); } }

        public AsyncCommand Roll { get; private set; }

        public int Sum { get => _sum; set { Set(nameof(Sum), ref _sum, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            History = await DB.DatabaseWrapper.GetCollectionFromQueryAsync<DiceRoll, ObservableCollection<DiceRoll>>(DbDiceRollsQuery).ConfigureAwait(true);
            if (History.Any())
            {
                NumberOfDices = History.First().Dices;
                NumberofSides = History.First().Sides;
            }
            IsBusy = false;
        }

        private async Task RollAsync()
        {
            Results.Clear();
            for (int i = 0; i < NumberOfDices; i++)
            {
                var rng = StandardRng.Create();
                Random random = RandomShim.Create(rng);
                Results.Add(random.Next(0, NumberofSides + 1));
            }
            RaisePropertyChanged(nameof(Results));
            Sum = Results.Sum(x => x);
            var history = new DiceRoll()
            {
                Dices = NumberOfDices,
                Sides = NumberofSides,
                Sum = Sum,
                Results = await GetResulsStringAsync().ConfigureAwait(true)
            };
            var task = Task.Run(() =>
            {
                using var charactersDb = DB.DatabaseWrapper.CreateCharactersDb();
                return charactersDb.Insert(history);
            });
            await task.ConfigureAwait(true);
            History.Insert(0, history);
        }

        private Task<string> GetResulsStringAsync()
        {
            var result = Task.Run(() =>
            {
                var sb = new StringBuilder();
                for (int i = 0; i < Results.Count; i++)
                {
                    int item = Results[i];
                    sb.Append(item);
                    if (i < Results.Count - 1)
                    {
                        sb.Append(", ");
                    }
                }
                return sb.ToString();
            });
            return result;
        }
    }
}