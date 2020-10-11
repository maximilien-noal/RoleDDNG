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
using RoleDDNG.ViewModels.Menus.Rules;
using RoleDDNG.ViewModels.Menus.Tools;

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

        private bool _isStartingUp = true;

        private ObservableCollection<IDocumentViewModel> _items = new ObservableCollection<IDocumentViewModel>();

        public MainViewModel()
        {
            ChangeBackgroundCommand = new RelayCommand(ChangeBackgroundMethod);
            ShowDiceRollWindow = new RelayCommand(() => AddDocumentViewModel<DiceRollViewModel>());
            ShowCharactersXpWindow = new AsyncCommand(async () => await OpenCharacterDbDependantViewAsync<CharactersXpViewModel>().ConfigureAwait(false));
            ShowTownGeneratorWindow = new RelayCommand(() => AddDocumentViewModel<TownGeneratorViewModel>());
            OpenCharactersDataBase = new AsyncCommand(async () => await AskAndOpenCharacterDbFileAsync().ConfigureAwait(false));
            OpenCharacterSheet = new AsyncCommand(async () => await OpenCharacterDbDependantViewAsync<OpenCharacterViewModel>().ConfigureAwait(false));
            OpenRacesDescriptions = new AsyncCommand(async () => await OpenCharacterDbDependantViewAsync<RacesDescriptionsViewModel>().ConfigureAwait(false));
            OpenSpellsDescriptions = new AsyncCommand(async () => await OpenCharacterDbDependantViewAsync<SpellsDescriptionsViewModel>().ConfigureAwait(false));
            OpenDonsDescriptions = new AsyncCommand(async () => await OpenCharacterDbDependantViewAsync<DonsDescriptionsViewModel>().ConfigureAwait(false));
            ChangeBackgroundCommand.Execute(this);
        }

        private void ChangeBackgroundMethod() => BackgroundSource = SimpleIoc.Default.GetInstance<IBackgroundSource>().GetBackgroundSource(BackgroundSource);

        public string BackgroundSource { get => _backgroundSource; private set { Set(nameof(BackgroundSource), ref _backgroundSource, value); } }

        public string CurrentCharacterDb { get => _currentCharacterDb; private set { Set(nameof(CurrentCharacterDb), ref _currentCharacterDb, value); } }

        public bool IsBusy { get => _isBusy; private set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public ObservableCollection<IDocumentViewModel> Items { get => _items; private set { Set(nameof(Items), ref _items, value); } }

        public AsyncCommand OpenCharactersDataBase { get; private set; }

        public AsyncCommand OpenCharacterSheet { get; private set; }

        public AsyncCommand OpenDonsDescriptions { get; }

        public RelayCommand ChangeBackgroundCommand { get; }

        public AsyncCommand OpenRacesDescriptions { get; private set; }

        public AsyncCommand OpenSpellsDescriptions { get; private set; }

        public AsyncCommand ShowCharactersXpWindow { get; private set; }

        public RelayCommand ShowDiceRollWindow { get; private set; }

        public RelayCommand ShowTownGeneratorWindow { get; private set; }

        public async Task ExitAppAsync()
        {
            await SaveAppSettingsAsync().ConfigureAwait(false);
        }

        public async Task<bool> LoadAppSettingsAsync()
        {
            IsBusy = true;
            var foundCharacterDb = false;
            if (File.Exists(_appSettingsFilePath))
            {
                var serializer = SimpleIoc.Default.GetInstance<IAsyncSerializer<AppSettings>>();
                var appSettings = await serializer.DeserializeAsync<AppSettings>(_appSettingsFilePath).ConfigureAwait(false);
                SimpleIoc.Default.Register(() => appSettings);
                foundCharacterDb = await OpenCharactersDatabaseAsync().ConfigureAwait(false);
            }
            else
            {
                SimpleIoc.Default.Register(() => new AppSettings());
            }
            IsBusy = false;
            _isStartingUp = false;
            if (!foundCharacterDb)
            {
                SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath = "";
            }
            return foundCharacterDb;
        }

        public void RemoveDocumentViewModel<T>() where T : IDocumentViewModel
        {
            if (Items.OfType<T>().Any())
            {
                Items.Remove(Items.OfType<T>().First());
            }
        }

        private static async Task<string> AskForCharactersDbFileAsync()
        {
            var fileDialog = SimpleIoc.Default.GetInstance<IFileDialog>();
            var dbFile = await fileDialog.OpenFileDialogAsync("Ouvrir une base de données de personnages...", "mdb").ConfigureAwait(false);
            return dbFile;
        }

        private void AddDocumentViewModel<T>() where T : IDocumentViewModel, new()
        {
            if (!Items.OfType<T>().Any())
            {
                var viewModel = new T();
                Items.Add(viewModel);
            }
        }

        private async Task<bool> AskAndOpenCharacterDbFileAsync()
        {
            var dbFile = await AskForCharactersDbFileAsync().ConfigureAwait(false);
            return await OpenCharacterDbAsync(dbFile).ConfigureAwait(false);
        }

        private async Task<bool> OpenCharacterDbAsync(string dbFile)
        {
            if (File.Exists(dbFile) &&
                await new AccessDb(dbFile).CanConnectAsync().ConfigureAwait(false))
            {
                SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath = dbFile;
                CurrentCharacterDb = dbFile;
                await SaveAppSettingsAsync().ConfigureAwait(false);
                return true;
            }
            return false;
        }

        private async Task OpenCharacterDbDependantViewAsync<T>() where T : IDocumentViewModel, new()
        {
            if (await OpenCharacterDbIfNoneOpenAsync().ConfigureAwait(false))
            {
                AddDocumentViewModel<T>();
            }
        }

        private async Task<bool> OpenCharacterDbIfNoneOpenAsync()
        {
            if (File.Exists(SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath))
            {
                return true;
            }
            return await OpenCharactersDatabaseAsync().ConfigureAwait(false);
        }

        private async Task<bool> OpenCharactersDatabaseAsync()
        {
            var dbFile = SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath;
            if (!File.Exists(dbFile))
            {
                if (_isStartingUp)
                {
                    return false;
                }
                dbFile = await AskForCharactersDbFileAsync().ConfigureAwait(false);
            }
            return await OpenCharacterDbAsync(dbFile).ConfigureAwait(false);
        }

        private async Task SaveAppSettingsAsync()
        {
            IsBusy = true;
            var configDir = Path.GetDirectoryName(_appSettingsFilePath);
            Directory.CreateDirectory(configDir);
            await SimpleIoc.Default.GetInstance<IAsyncSerializer<AppSettings>>().SerializeAsync(_appSettingsFilePath, SimpleIoc.Default.GetInstance<AppSettings>()).ConfigureAwait(false);
            IsBusy = false;
        }
    }
}