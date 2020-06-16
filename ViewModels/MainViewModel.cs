using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.Interfaces.Backgrounds;
using RoleDDNG.Interfaces.Serialization;
using RoleDDNG.Models.Options;
using RoleDDNG.ViewModels.Interfaces;
using RoleDDNG.ViewModels.ToolsVMs;

using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels
{
    public sealed class MainViewModel : ViewModelBase
    {
        private readonly string _appSettingsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RoleDDNG\\ROLE.CFG");

        private AppSettings _appSettings = new AppSettings();

        private string _backgroundSource = string.Empty;

        private bool _isBusy = true;

        private ObservableCollection<IContent> _items = new ObservableCollection<IContent>();

        private IContent? _selectedWindow;

        public MainViewModel()
        {
            SetCharacterDatabasePath = new RelayCommand<string>(SetCharacterDatabasePathMethod);
            ShowDiceRollWindow = new RelayCommand(() => AddMdiWindow<DiceRollViewModel>());
            ShowTownGeneratorWindow = new RelayCommand(() => AddMdiWindow<TownGeneratorViewModel>());
            OpenCharactersDataBase = new AsyncCommand(OpenCharactersDataBaseAsync);
            BackgroundSource = SimpleIoc.Default.GetInstance<IBackgroundSource>().GetBackgroundSource();
        }

        public AppSettings AppSettings { get => _appSettings; private set { Set(nameof(AppSettings), ref _appSettings, value); } }

        public string BackgroundSource { get => _backgroundSource; private set { Set(nameof(BackgroundSource), ref _backgroundSource, value); } }

        public bool IsBusy { get => _isBusy; private set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public ObservableCollection<IContent> Items { get => _items; private set { Set(nameof(Items), ref _items, value); } }

        public AsyncCommand OpenCharactersDataBase { get; private set; }

        public IContent? SelectedWindow { get => _selectedWindow; set { Set(nameof(SelectedWindow), ref _selectedWindow, value); } }

        public RelayCommand<string> SetCharacterDatabasePath { get; private set; }

        public RelayCommand ShowDiceRollWindow { get; private set; }

        public RelayCommand ShowTownGeneratorWindow { get; private set; }

        public async Task ExitAppAsync()
        {
            IsBusy = true;
            string configDir = Path.GetDirectoryName(_appSettingsFilePath);
            if (!Directory.Exists(configDir))
            {
                Directory.CreateDirectory(configDir);
            }

            await SimpleIoc.Default.GetInstance<IAsyncSerializer<AppSettings>>().SerializeAsync(_appSettingsFilePath, AppSettings).ConfigureAwait(false);
            IsBusy = false;
        }

        public async Task LoadAppSettingsAsync()
        {
            IsBusy = true;
            if (File.Exists(_appSettingsFilePath))
            {
                var serializer = SimpleIoc.Default.GetInstance<IAsyncSerializer<AppSettings>>();
                AppSettings = await serializer.DeserializeAsync(_appSettingsFilePath).ConfigureAwait(false);
            }
            IsBusy = false;
        }

        public void RemoveMdiWindow<T>() where T : IContent
        {
            if (Items.OfType<T>().Any())
            {
                Items.Remove(Items.OfType<T>().First());
            }
        }

        public async Task ShowCharacterChoiceAsync()
        {
            if (!string.IsNullOrWhiteSpace(AppSettings.LastCharacterDBPath) && File.Exists(AppSettings.LastCharacterDBPath))
            {
                AddMdiWindow<OpenCharacterViewModel>();
                var characterDBViewModel = SimpleIoc.Default.GetInstance<OpenCharacterViewModel>();
                await characterDBViewModel.GetCharactersFromDbAsync(AppSettings.LastCharacterDBPath).ConfigureAwait(true);
            }
        }

        private void AddMdiWindow<T>() where T : IContent, new()
        {
            if (!Items.OfType<T>().Any())
            {
                var viewModel = new T();
                Items.Add(viewModel);
            }
        }

        private async Task OpenCharactersDataBaseAsync()
        {
            AddMdiWindow<OpenCharacterViewModel>();
            var characterDBViewModel = SimpleIoc.Default.GetInstance<OpenCharacterViewModel>();
            await characterDBViewModel.AskForDatabaseFileCommand.ExecuteAsync().ConfigureAwait(true);
        }

        private void SetCharacterDatabasePathMethod(string path)
        {
            AppSettings.LastCharacterDBPath = path;
        }

#pragma warning disable CA1822 // Static bindings work, but make the designer view throw an error.
    }

#pragma warning restore CA1822
}