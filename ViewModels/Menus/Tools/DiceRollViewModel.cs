using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using RoleDDNG.ViewModels.Interfaces;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using RandN.Compat;
using RandN;

namespace RoleDDNG.ViewModels.Menus.Tools
{
    public class DiceRollViewModel : ViewModelBase, IDocumentViewModel
    {
        private int _numberOfDices = 1;

        private int _numberOfSides = 6;

        private ObservableCollection<int> _results = new ObservableCollection<int>();

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
            Roll = new RelayCommand(Roll_Execute);
        }

        public List<int> DiceTypes { get; private set; } = new List<int>();

        public bool IsBusy => false;

        public int NumberOfDices { get => _numberOfDices; set { Set(nameof(NumberOfDices), ref _numberOfDices, value); } }

        public int NumberofSides { get => _numberOfSides; set { Set(nameof(NumberofSides), ref _numberOfSides, value); } }

        public ObservableCollection<int> Results { get => _results; private set { Set(nameof(Results), ref _results, value); } }

        public RelayCommand Roll { get; private set; }

        public int Sum { get => _sum; set { Set(nameof(Sum), ref _sum, value); } }

        private void Roll_Execute()
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
        }
    }
}