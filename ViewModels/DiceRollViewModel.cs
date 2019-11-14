using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using RoleDDNG.Models;
using RoleDDNG.ViewModels.Interfaces;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace RoleDDNG.ViewModels
{
    public class DiceRollViewModel : ViewModelBase, IContent
    {
        private static readonly Random _rng = new Random();

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
            CloseCommand = new RelayCommand(Close_Execute);
        }

        private void Close_Execute()
        {
            if (CanClose)
            {
                Closing?.Invoke(this, EventArgs.Empty);
            }
        }

        private void Roll_Execute()
        {
            Results.Clear();
            for (int i = 0; i < NumberOfDices; i++)
            {
                Results.Add(_rng.Next(0, NumberofSides + 1));
            }
            RaisePropertyChanged(nameof(Results));
        }

        public string Title => "Lanceur de dés";

        public bool CanClose => true;

        private int _numberOfDices = 1;

        public int NumberOfDices { get => _numberOfDices; set { Set(nameof(NumberOfDices), ref _numberOfDices, value); } }

        private int _numberOfSides = 6;

        public int NumberofSides { get => _numberOfSides; set { Set(nameof(NumberofSides), ref _numberOfSides, value); } }

        public event EventHandler Closing;

        private ObservableCollection<int> _results = new ObservableCollection<int>();
        public ObservableCollection<int> Results { get => _results; private set { Set(nameof(Results), ref _results, value); } }

        public RelayCommand Roll { get; private set; }

        public RelayCommand CloseCommand { get; private set; }

        public List<int> DiceTypes { get; private set; } = new List<int>();
    }
}