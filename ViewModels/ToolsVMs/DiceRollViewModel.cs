using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;

using RoleDDNG.ViewModels.Interfaces;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace RoleDDNG.ViewModels.ToolsVMs
{
    public class DiceRollViewModel : ViewModelBase, IContent
    {
        private static readonly Random _rng = new Random(1000);

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

        private void Roll_Execute()
        {
            Results.Clear();
            for (int i = 0; i < NumberOfDices; i++)
            {
                Results.Add(_rng.Next(0, NumberofSides + 1));
            }
            RaisePropertyChanged(nameof(Results));
            Sum = Results.Sum(x => x);
        }

        public string Title => "Lanceur de dés";

        private int _numberOfDices = 1;

        public int NumberOfDices { get => _numberOfDices; set { Set(nameof(NumberOfDices), ref _numberOfDices, value); } }

        private int _numberOfSides = 6;

        public int NumberofSides { get => _numberOfSides; set { Set(nameof(NumberofSides), ref _numberOfSides, value); } }

        private int _sum = 0;

        public int Sum { get => _sum; set { Set(nameof(Sum), ref _sum, value); } }

        private ObservableCollection<int> _results = new ObservableCollection<int>();
        public ObservableCollection<int> Results { get => _results; private set { Set(nameof(Results), ref _results, value); } }

        public RelayCommand Roll { get; private set; }

        public List<int> DiceTypes { get; private set; } = new List<int>();
    }
}