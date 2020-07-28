using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.DatabaseLayer;
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

        private string _backgroundSource = string.Empty;

        private string _currentCharacterDb = "";

        private bool _isBusy = true;

        private ObservableCollection<IDocumentViewModel> _items = new ObservableCollection<IDocumentViewModel>();

        private IDocumentViewModel? _selectedWindow;

        public MainViewModel()
        {
            ShowDiceRollWindow = new RelayCommand(() => AddMdiWindow<DiceRollViewModel>());
            ShowCharactersXpWindow = new AsyncCommand(async () => await OpenCharacterDbDependantViewAsync<CharactersXpViewModel>().ConfigureAwait(true));
            ShowTownGeneratorWindow = new RelayCommand(() => AddMdiWindow<TownGeneratorViewModel>());
            OpenCharactersDataBase = new AsyncCommand(async () => await AskAndOpenCharacterDbFileAsync().ConfigureAwait(true));
            OpenCharacterSheet = new AsyncCommand(async () => await OpenCharacterDbDependantViewAsync<OpenCharacterViewModel>().ConfigureAwait(true));
            BackgroundSource = SimpleIoc.Default.GetInstance<IBackgroundSource>().GetBackgroundSource();
        }

        public string BackgroundSource { get => _backgroundSource; private set { Set(nameof(BackgroundSource), ref _backgroundSource, value); } }

        public string CurrentCharacterDb { get => _currentCharacterDb; private set { Set(nameof(CurrentCharacterDb), ref _currentCharacterDb, value); } }

        public bool IsBusy { get => _isBusy; private set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public ObservableCollection<IDocumentViewModel> Items { get => _items; private set { Set(nameof(Items), ref _items, value); } }

        public AsyncCommand OpenCharactersDataBase { get; private set; }

        public AsyncCommand OpenCharacterSheet { get; private set; }

        public IDocumentViewModel? SelectedWindow { get => _selectedWindow; set { Set(nameof(SelectedWindow), ref _selectedWindow, value); } }

        public AsyncCommand ShowCharactersXpWindow { get; private set; }

        public RelayCommand ShowDiceRollWindow { get; private set; }

        public RelayCommand ShowTownGeneratorWindow { get; private set; }

        public async Task ExitAppAsync()
        {
            await SaveAppSettingsAsync().ConfigureAwait(true);
        }

        public async Task<bool> LoadAppSettingsAsync()
        {
            IsBusy = true;
            if (File.Exists(_appSettingsFilePath))
            {
                var serializer = SimpleIoc.Default.GetInstance<IAsyncSerializer<AppSettings>>();
                var appSettings = await serializer.DeserializeAsync<AppSettings>(_appSettingsFilePath).ConfigureAwait(true);
                SimpleIoc.Default.Register<AppSettings>(() => appSettings);
                return await OpenCharactersDatabaseAsync().ConfigureAwait(true);
            }
            IsBusy = false;
            return false;
        }

        public void RemoveMdiWindow<T>() where T : IDocumentViewModel
        {
            if (Items.OfType<T>().Any())
            {
                Items.Remove(Items.OfType<T>().First());
            }
        }

        private static async Task<string> AskForCharactersDbFileAsync()
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

        private async Task<bool> AskAndOpenCharacterDbFileAsync()
        {
            var dbFile = await AskForCharactersDbFileAsync().ConfigureAwait(true);
            return await OpenCharacterDbAsync(dbFile).ConfigureAwait(true);
        }

        private async Task<bool> OpenCharacterDbAsync(string dbFile)
        {
            if (!string.IsNullOrWhiteSpace(dbFile) &&
                File.Exists(dbFile) &&
                await new AccessDb(dbFile).CanConnectAsync().ConfigureAwait(true))
            {
                SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath = dbFile;
                CurrentCharacterDb = dbFile;
                await SaveAppSettingsAsync().ConfigureAwait(true);
                return true;
            }
            return false;
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
            if (!string.IsNullOrWhiteSpace(SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath) &&
                File.Exists(SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath))
            {
                return true;
            }
            return await OpenCharactersDatabaseAsync().ConfigureAwait(true);
        }

        private async Task<bool> OpenCharactersDatabaseAsync()
        {
            var dbFile = SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath;
            if (string.IsNullOrWhiteSpace(dbFile))
            {
                dbFile = await AskForCharactersDbFileAsync().ConfigureAwait(true);
            }
            return await OpenCharacterDbAsync(dbFile).ConfigureAwait(true);
        }

        private async Task SaveAppSettingsAsync()
        {
            IsBusy = true;
            var configDir = Path.GetDirectoryName(_appSettingsFilePath);
            if (!Directory.Exists(configDir))
            {
                Directory.CreateDirectory(configDir);
            }

            await SimpleIoc.Default.GetInstance<IAsyncSerializer<AppSettings>>().SerializeAsync(_appSettingsFilePath, SimpleIoc.Default.GetInstance<AppSettings>()).ConfigureAwait(false);
            IsBusy = false;
        }
    }
}