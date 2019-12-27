using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;

using Microsoft.VisualStudio.Threading;

using RoleDDNG.Interfaces.Serialization;
using RoleDDNG.Models.Options;
using RoleDDNG.ViewModels.Interfaces;
using RoleDDNG.ViewModels.RNG;
using RoleDDNG.ViewModels.ToolsVMs;

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

        private readonly IAsyncSerializer<AppSettings> _settingsSerializer;

        public MainViewModel(IAsyncSerializer<AppSettings> serializer)
        {
            IsBusy = true;

            ShowDiceRollWindow = new RelayCommand(ShowDiceRollWindowMethod);
            ShowTownGeneratorWindow = new RelayCommand(ShowTownGeneratorWindowMethod);
            ExitApp = new AsyncCommand(ExitAppMethodAsync);
            LoadAppSettings = new AsyncCommand(LoadAppSettingsMethodAsync);
            _settingsSerializer = serializer ?? throw new ArgumentNullException(nameof(serializer));

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
            LoadAppSettings.Execute(this);
            IsBusy = false;
        }

        public IAsyncCommand LoadAppSettings { get; private set; }

        private async Task LoadAppSettingsMethodAsync()
        {
            if (File.Exists(_appSettingsFilePath))
            {
                AppSettings = await _settingsSerializer.DeserializeAsync(_appSettingsFilePath).ConfigureAwait(true);
            }
        }

        private void ShowTownGeneratorWindowMethod()
        {
            AddMdiWindow<TownGeneratorViewModel>();
        }

        private void ShowDiceRollWindowMethod()
        {
            AddMdiWindow<DiceRollViewModel>();
        }

        private void AddMdiWindow<T>() where T : IContent, new()
        {
            if (Items.OfType<T>().Any() == false)
            {
                var viewModel = new T();
                Items.Add(viewModel);
            }
        }

        private readonly List<string> _backgrounds = new List<string>();

        private string _backgroundSource = "Assets/backgrounds/original.png";

        public string BackgroundSource { get => _backgroundSource; private set { Set(nameof(BackgroundSource), ref _backgroundSource, value); } }

        private ObservableCollection<IContent> _items = new ObservableCollection<IContent>();

        public ObservableCollection<IContent> Items { get => _items; private set { Set(nameof(Items), ref _items, value); } }

        public RelayCommand ShowDiceRollWindow { get; private set; }
        public RelayCommand ShowTownGeneratorWindow { get; private set; }

#pragma warning disable CA1822 // Static bindings work, but make the designer view throw an error.
        public IAsyncCommand ExitApp { get; private set; }

        private async Task ExitAppMethodAsync()
        {
            IsBusy = true;
            if (Directory.Exists(Path.GetDirectoryName(_appSettingsFilePath)) == false)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_appSettingsFilePath));
            }

            await _settingsSerializer.SerializeAsync(_appSettingsFilePath, AppSettings).ConfigureAwait(false);
            IsBusy = false;
        }
    }

#pragma warning restore CA1822
}