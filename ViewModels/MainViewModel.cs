using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using Microsoft.VisualStudio.Threading;
using RoleDDNG.Interfaces.Serialization;
using RoleDDNG.Models.Options;
using RoleDDNG.Models.Structs;
using RoleDDNG.ViewModels.Interfaces;
using RoleDDNG.ViewModels.RNG;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels
{
    public sealed class MainViewModel : ViewModelBase, IBusyStateNotifier
    {
        private bool _isBusy = true;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        private IContent _selectedWindow = default;
        public IContent SelectedWindow { get => _selectedWindow; set { Set(nameof(SelectedWindow), ref _selectedWindow, value); } }

        private readonly string _appSettingsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RoleDDNG\\ROLE.CFG");

        private AppSettings _appSettings = new AppSettings();

        public AppSettings AppSettings { get => _appSettings; set { Set(nameof(AppSettings), ref _appSettings, value); } }

        private readonly ISerializer<AppSettings> _settingsSerializer;

        public MainViewModel(ISerializer<AppSettings> serializer)
        {
            IsBusy = true;
            if (Directory.Exists(Path.GetDirectoryName(_appSettingsFilePath)) == false)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_appSettingsFilePath));
            }
            _settingsSerializer = serializer ?? throw new ArgumentNullException(nameof(serializer));
            if (File.Exists(_appSettingsFilePath))
            {
                AppSettings = _settingsSerializer.Deserialize(_appSettingsFilePath);
            }
            if (DateTime.Now.Month == 12)
            {
                _backgrounds.Add("Assets/backgrounds/christmas.jpg");
                _backgrounds.Add("Assets/backgrounds/christmas2.png");
                _backgrounds.Add("Assets/backgrounds/christmas3.jpg");
                _backgrounds.Add("Assets/backgrounds/christmasEve.jpg");
                _backgrounds.Add("Assets/backgrounds/IceQueen.png");
                _backgrounds.Add("Assets/backgrounds/imp.png");
            }
            else if (DateTime.Now.Month == 10)
            {
                _backgrounds.Add("Assets/backgrounds/sorceress.png");
                _backgrounds.Add("Assets/backgrounds/halloween.jpg");
                _backgrounds.Add("Assets/backgrounds/halloween2.png");
                _backgrounds.Add("Assets/backgrounds/halloween3.png");
                _backgrounds.Add("Assets/backgrounds/halloween4.jpg");
                _backgrounds.Add("Assets/backgrounds/halloween5.png");
                _backgrounds.Add("Assets/backgrounds/halloween6.png");
                _backgrounds.Add("Assets/backgrounds/halloween7.png");
                _backgrounds.Add("Assets/backgrounds/halloween8.png");
            }
            else
            {
                _backgrounds.Add("Assets/backgrounds/citiesInTheSky.png");
                _backgrounds.Add("Assets/backgrounds/deer.jpg");
                _backgrounds.Add("Assets/backgrounds/deer2.jpg");
                _backgrounds.Add("Assets/backgrounds/diceTable.jpg");
                _backgrounds.Add("Assets/backgrounds/dwarfHouse.jpg");
                _backgrounds.Add("Assets/backgrounds/faery.png");
                _backgrounds.Add("Assets/backgrounds/forest.png");
                _backgrounds.Add("Assets/backgrounds/forest2.jpg");
                _backgrounds.Add("Assets/backgrounds/LotsOfDices.png");
                _backgrounds.Add("Assets/backgrounds/merchant.jpg");
                _backgrounds.Add("Assets/backgrounds/original.png");
                _backgrounds.Add("Assets/backgrounds/sword.png");
                _backgrounds.Add("Assets/backgrounds/unicorn.png");
                _backgrounds.Add("Assets/backgrounds/warrior.jpg");
            }
            BackgroundSource = _backgrounds[StaticRNG.RNG.Next(0, _backgrounds.Count)];
            ShowDiceRollWindow = new RelayCommand(ShowDiceRollWindow_Execute);
            ExitApp = new RelayCommand(ExitApp_Execute);
            IsBusy = false;
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

        private readonly List<string> _backgrounds = new List<string>();

        private string _backgroundSource = "Assets/backgrounds/original.png";

        public string BackgroundSource { get => _backgroundSource; private set { Set(nameof(BackgroundSource), ref _backgroundSource, value); } }

        private ObservableCollection<IContent> _items = new ObservableCollection<IContent>();

        public ObservableCollection<IContent> Items { get => _items; private set { Set(nameof(Items), ref _items, value); } }

        public RelayCommand ShowDiceRollWindow { get; private set; }

#pragma warning disable CA1822 // Static bindings work, but make the designer view throw an error.
        public RelayCommand ExitApp { get; private set; }

        private void ExitApp_Execute()
        {
            IsBusy = true;
            _settingsSerializer.Serialize(_appSettingsFilePath, AppSettings);
            IsBusy = false;
        }
    }

#pragma warning restore CA1822
}