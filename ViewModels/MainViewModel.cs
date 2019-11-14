using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using RoleDDNG.ViewModels.Interfaces;
using System;
using System.Linq;
using System.Collections.ObjectModel;
using System.Windows;

namespace RoleDDNG.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private IContent _selectedWindow = null;
        public IContent SelectedWindow { get => _selectedWindow; set { Set(nameof(SelectedWindow), ref _selectedWindow, value); } }

        public MainViewModel()
        {
            ShowDiceRollWindow = new RelayCommand(ShowDiceRollWindow_Execute);
        }

        private void ShowDiceRollWindow_Execute()
        {
            if (Items.OfType<DiceRollViewModel>().Any() == false)
            {
                var diceViewModel = new DiceRollViewModel();
                diceViewModel.Closing += delegate { Items.Remove(diceViewModel); };
                Items.Add(diceViewModel);
            }
        }

        private ObservableCollection<IContent> _items = new ObservableCollection<IContent>();

        public ObservableCollection<IContent> Items { get => _items; private set { Set(nameof(Items), ref _items, value); } }

        public RelayCommand ShowDiceRollWindow { get; private set; }

#pragma warning disable CA1822 // Static bindings work, but makes the designer view throw an error.
        public RelayCommand ExitAppCommand { get => new RelayCommand(() => { Application.Current.Dispatcher.Invoke(() => Application.Current.MainWindow.Close()); }); }
    }

#pragma warning restore CA1822
}