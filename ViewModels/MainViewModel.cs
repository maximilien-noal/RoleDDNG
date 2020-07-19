using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.DatabaseLayer;
using RoleDDNG.DatabaseLayer.Models;
using RoleDDNG.DatabaseLayer.Enums;
using RoleDDNG.Interfaces.Backgrounds;
using RoleDDNG.Interfaces.Dialogs;
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

        private ObservableCollection<IDocumentViewModel> _items = new ObservableCollection<IDocumentViewModel>();

        private IDocumentViewModel? _selectedWindow;

        public MainViewModel()
        {
            ShowDiceRollWindow = new RelayCommand(() => AddMdiWindow<DiceRollViewModel>());
            ShowCharactersXpWindow = new AsyncCommand(async () => await OpenCharacterDbDependantViewAsync<CharactersXpViewModel>().ConfigureAwait(true));
            ShowTownGeneratorWindow = new RelayCommand(() => AddMdiWindow<TownGeneratorViewModel>());
            OpenCharactersDataBase = new AsyncCommand(OpenCharactersDatabaseAsync);
            OpenCharacterSheet = new AsyncCommand(async () => await OpenCharacterDbDependantViewAsync<OpenCharacterViewModel>().ConfigureAwait(true));
            BackgroundSource = SimpleIoc.Default.GetInstance<IBackgroundSource>().GetBackgroundSource();
        }

        public AppSettings AppSettings { get => _appSettings; private set { Set(nameof(AppSettings), ref _appSettings, value); } }

        public string BackgroundSource { get => _backgroundSource; private set { Set(nameof(BackgroundSource), ref _backgroundSource, value); } }

        public bool IsBusy { get => _isBusy; private set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public ObservableCollection<IDocumentViewModel> Items { get => _items; private set { Set(nameof(Items), ref _items, value); } }

        public AsyncCommand OpenCharactersDataBase { get; private set; }

        public AsyncCommand OpenCharacterSheet { get; private set; }

        public IDocumentViewModel? SelectedWindow { get => _selectedWindow; set { Set(nameof(SelectedWindow), ref _selectedWindow, value); } }

        public AsyncCommand ShowCharactersXpWindow { get; private set; }

        public RelayCommand ShowDiceRollWindow { get; private set; }

        public RelayCommand ShowTownGeneratorWindow { get; private set; }

        public static async Task<bool> CheckIfCharactersDbFileIsValidAsync(string dbFile)
        {
            return !string.IsNullOrWhiteSpace(dbFile) &&
                            File.Exists(dbFile) &&
                            await new DbAccessor(new Database(dbFile, DbType.UserCharactersDb)).CanConnectAsync().ConfigureAwait(true);
        }

        public async Task ExitAppAsync()
        {
            IsBusy = true;
            var configDir = Path.GetDirectoryName(_appSettingsFilePath);
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
                AppSettings = await serializer.DeserializeAsync<AppSettings>(_appSettingsFilePath).ConfigureAwait(false);
            }
            IsBusy = false;
        }

        public async Task OpenCharactersDatabaseAsync()
        {
            string dbFile = await OpenCharactersDbFileDialogAsync().ConfigureAwait(true);
            if (!string.IsNullOrWhiteSpace(dbFile) &&
                File.Exists(dbFile) &&
                await new DbAccessor(new Database(dbFile, DbType.UserCharactersDb)).CanConnectAsync().ConfigureAwait(true))
            {
                AppSettings.LastCharacterDBPath = dbFile;
                await OpenCharacterSheet.ExecuteAsync().ConfigureAwait(true);
            }
        }

        public void RemoveMdiWindow<T>() where T : IDocumentViewModel
        {
            if (Items.OfType<T>().Any())
            {
                Items.Remove(Items.OfType<T>().First());
            }
        }

        private static async Task<string> OpenCharactersDbFileDialogAsync()
        {
            var fileDialog = SimpleIoc.Default.GetInstance<IFileDialog>();
            var dbFile = await fileDialog.OpenFileDialogAsync("Ouvrir une base de données de personnages...", "mdb").ConfigureAwait(true);
            return dbFile;
        }

        private void AddMdiWindow<T>() where T : IDocumentViewModel, new()
        {
            if (!Items.OfType<T>().Any())
            {
                var viewModel = new T();
                Items.Add(viewModel);
                return;
            }
            var existingViewModel = Items.OfType<T>().FirstOrDefault();
            if (existingViewModel != null)
            {
                Items.Remove(existingViewModel);
                Items.Add(existingViewModel);
            }
        }

        private async Task OpenCharacterDbDependantViewAsync<T>() where T : IDocumentViewModel, new()
        {
            if (await OpenCharacterDbIfNoneOpenAsync().ConfigureAwait(true))
            {
                AddMdiWindow<T>();
            }
        }

        private async Task<bool> OpenCharacterDbIfNoneOpenAsync()
        {
            if (!string.IsNullOrWhiteSpace(AppSettings.LastCharacterDBPath) &&
                File.Exists(AppSettings.LastCharacterDBPath))
            {
                return true;
            }
            string dbFile = await OpenCharactersDbFileDialogAsync().ConfigureAwait(true);
            if (await CheckIfCharactersDbFileIsValidAsync(dbFile).ConfigureAwait(true))
            {
                AppSettings.LastCharacterDBPath = dbFile;
                return true;
            }
            return false;
        }

#pragma warning disable CA1822 // Static bindings work, but make the designer view throw an error.
    }

#pragma warning restore CA1822
}