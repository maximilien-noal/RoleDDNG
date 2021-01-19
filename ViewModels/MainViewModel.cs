using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.DatabaseLayer;
using RoleDDNG.Interfaces.Backgrounds;
using RoleDDNG.Interfaces.Dialogs;
using RoleDDNG.Interfaces.Serialization;
using RoleDDNG.Models.Options;
using RoleDDNG.ViewModels.DB;
using RoleDDNG.ViewModels.Interfaces;
using RoleDDNG.ViewModels.Menus.Characters;
using RoleDDNG.ViewModels.Menus.Rules;
using RoleDDNG.ViewModels.Menus.Tools;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace RoleDDNG.ViewModels
{
    public sealed class MainViewModel : ViewModelBase
    {
        private readonly string _appSettingsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RoleDDNG\\ROLE.CFG");

        private string _backgroundSource = string.Empty;

        private string _currentCharacterDb = "";

        private bool _isBusy = true;

        private bool _isStartingUp = true;

        private ObservableCollection<dynamic> _items = new ObservableCollection<dynamic>();

        public MainViewModel()
        {
            OpenCharacterImport = new AsyncCommand(async () => await OpenCharacterImportMethodAsync().ConfigureAwait(true));
            ChangeBackgroundCommand = new RelayCommand(ChangeBackgroundMethod);
            ShowDiceRollWindow = new AsyncCommand(async () => await OpenDbDependantViewModelAsync<DiceRollViewModel>().ConfigureAwait(true));
            ShowCharactersXpWindow = new AsyncCommand(async () => await OpenDbDependantViewModelAsync<CharactersXpViewModel>().ConfigureAwait(true));
            ShowTownGeneratorWindow = new RelayCommand(() => AddDocumentViewModel<TownGeneratorViewModel>());
            OpenCharactersDataBase = new AsyncCommand(async () => await AskAndOpenCharacterDbFileAsync().ConfigureAwait(true));
            OpenCharacterSheet = new AsyncCommand(async () => await OpenDbDependantViewModelAsync<OpenCharacterViewModel>().ConfigureAwait(true));
            OpenRacesDescriptions = new AsyncCommand(async () => await OpenDbDependantViewModelAsync<RacesDescriptionsViewModel>().ConfigureAwait(true));
            OpenSpellsDescriptions = new AsyncCommand(async () => await OpenDbDependantViewModelAsync<SpellsDescriptionsViewModel>().ConfigureAwait(true));
            OpenDonsDescriptions = new AsyncCommand(async () => await OpenDbDependantViewModelAsync<DonsDescriptionsViewModel>().ConfigureAwait(true));
            ChangeBackgroundCommand.Execute(this);
        }

        private async Task OpenCharacterImportMethodAsync()
        {
            if (await OpenCharacterDbIfNoneOpenAndRunMigrationsAsync().ConfigureAwait(true) == false)
            {
                return;
            }
            var otherCharactersDatabases = await OpenForeignCharacterDBsAsync().ConfigureAwait(true);
            if (otherCharactersDatabases.Any() == false)
            {
                return;
            }
            if (otherCharactersDatabases.Any(x => Path.GetFullPath(x).ToUpperInvariant() == Path.GetFullPath(CurrentCharacterDb).ToUpperInvariant()) && otherCharactersDatabases.Count() == 1)
            {
                MessageBox.Show("Ceci est la base de données de personnages déjà ouverte.");
                return;
            }
            if (otherCharactersDatabases.Any(x => Path.GetFullPath(x).ToUpperInvariant() == DB.DatabaseWrapper.GetFullProgDbPath().ToUpperInvariant()) && otherCharactersDatabases.Count() == 1)
            {
                MessageBox.Show("Ceci est la base de données du programme.");
                return;
            }
            otherCharactersDatabases = otherCharactersDatabases.Where(x => Path.GetFullPath(x).ToUpperInvariant() != Path.GetFullPath(CurrentCharacterDb).ToUpperInvariant() && Path.GetFullPath(x).ToUpperInvariant() != DB.DatabaseWrapper.GetFullProgDbPath().ToUpperInvariant());
            var viewModel = new CharacterImportViewModel(otherCharactersDatabases);
            Items.Add(viewModel);
        }

        private void ChangeBackgroundMethod() => BackgroundSource = SimpleIoc.Default.GetInstance<IBackgroundSource>().GetBackgroundSource(BackgroundSource);

        public string BackgroundSource { get => _backgroundSource; private set { Set(nameof(BackgroundSource), ref _backgroundSource, value); } }

        public string CurrentCharacterDb { get => _currentCharacterDb; private set { Set(nameof(CurrentCharacterDb), ref _currentCharacterDb, value); } }

        public bool IsBusy { get => _isBusy; private set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public ObservableCollection<object> Items { get => _items; private set { Set(nameof(Items), ref _items, value); } }

        public AsyncCommand OpenCharacterImport { get; private set; }

        public AsyncCommand OpenCharactersDataBase { get; private set; }

        public AsyncCommand OpenCharacterSheet { get; private set; }

        public AsyncCommand OpenDonsDescriptions { get; }

        public RelayCommand ChangeBackgroundCommand { get; }

        public AsyncCommand OpenRacesDescriptions { get; private set; }

        public AsyncCommand OpenSpellsDescriptions { get; private set; }

        public AsyncCommand ShowCharactersXpWindow { get; private set; }

        public AsyncCommand ShowDiceRollWindow { get; private set; }

        public RelayCommand ShowTownGeneratorWindow { get; private set; }

        public async Task ExitAppAsync()
        {
            await SaveAppSettingsAsync().ConfigureAwait(true);
        }

        public async Task<bool> LoadAppSettingsAsync()
        {
            IsBusy = true;
            var foundCharacterDb = false;
            if (File.Exists(_appSettingsFilePath))
            {
                var serializer = SimpleIoc.Default.GetInstance<IAsyncSerializer<AppSettings>>();
                var appSettings = await serializer.DeserializeAsync<AppSettings>(_appSettingsFilePath).ConfigureAwait(true);
                SimpleIoc.Default.Register(() => appSettings);
                foundCharacterDb = await OpenCharactersDatabaseAsync().ConfigureAwait(true);
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

        public void RemoveDocumentViewModel(dynamic viewModel)
        {
            var itemToRemove = Items.FirstOrDefault(x => viewModel.GetType() == x.GetType());
            if (itemToRemove != null)
            {
                Items.Remove(itemToRemove);
            }
        }

        private static async Task<string> AskForCharactersDbFileAsync(string title = "Ouvrir une base de données de personnages...")
        {
            var fileDialog = SimpleIoc.Default.GetInstance<IFileDialog>();
            var dbFile = await fileDialog.OpenFileDialogAsync(title, "mdb").ConfigureAwait(true);
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

        private static async Task<IEnumerable<string>> OpenForeignCharacterDBsAsync()
        {
            var fileDialog = SimpleIoc.Default.GetInstance<IFileDialog>();
            var dbFiles = await fileDialog.OpenFilesDialogAsync("Ouvrir une ou plusieurs bases de données de personnages à importer...", "mdb").ConfigureAwait(true);
            var validFiles = new List<string>();
            foreach (var dbFile in dbFiles)
            {
                if (File.Exists(dbFile) && await new AccessDb(dbFile).CanConnectAsync().ConfigureAwait(true))
                {
                    validFiles.Add(dbFile);
                }
            }
            return validFiles;
        }

        private async Task AskAndOpenCharacterDbFileAsync()
        {
            var dbFile = await AskForCharactersDbFileAsync().ConfigureAwait(true);
            await OpenCharacterDbAsync(dbFile).ConfigureAwait(true);
        }

        private async Task<bool> OpenCharacterDbAsync(string dbFile)
        {
            if (File.Exists(dbFile) &&
                await new AccessDb(dbFile).CanConnectAsync().ConfigureAwait(true))
            {
                SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath = dbFile;
                CurrentCharacterDb = dbFile;
                await SaveAppSettingsAsync().ConfigureAwait(true);
                return true;
            }
            return false;
        }

        private async Task OpenDbDependantViewModelAsync<T>() where T : IDocumentViewModel, new()
        {
            if (await OpenCharacterDbIfNoneOpenAndRunMigrationsAsync().ConfigureAwait(true))
            {
                AddDocumentViewModel<T>();
            }
        }

        private async Task<bool> OpenCharacterDbIfNoneOpenAndRunMigrationsAsync()
        {
            if (await OpenCharacterDbIfNoneOpenAsync().ConfigureAwait(true))
            {
                if (MigrationsRunner.NeedsToRun())
                {
                    IsBusy = true;
                    await new MigrationsRunner().RunCharactersDbMigrationsAsync().ConfigureAwait(true);
                    IsBusy = false;
                }
                return true;
            }
            return false;
        }

        private async Task<bool> OpenCharacterDbIfNoneOpenAsync()
        {
            if (File.Exists(SimpleIoc.Default.GetInstance<AppSettings>().LastCharacterDBPath))
            {
                return true;
            }
            return await OpenCharactersDatabaseAsync().ConfigureAwait(true);
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
                dbFile = await AskForCharactersDbFileAsync().ConfigureAwait(true);
            }
            return await OpenCharacterDbAsync(dbFile).ConfigureAwait(true);
        }

        private async Task SaveAppSettingsAsync()
        {
            IsBusy = true;
            var configDir = Path.GetDirectoryName(_appSettingsFilePath);
            if (configDir != null)
            {
                Directory.CreateDirectory(configDir);
            }
            await SimpleIoc.Default.GetInstance<IAsyncSerializer<AppSettings>>().SerializeAsync(_appSettingsFilePath, SimpleIoc.Default.GetInstance<AppSettings>()).ConfigureAwait(true);
            IsBusy = false;
        }
    }
}