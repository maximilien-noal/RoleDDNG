using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using Microsoft.VisualStudio.Threading;

using RoleDDNG.Interfaces.Backgrounds;
using RoleDDNG.Interfaces.Serialization;
using RoleDDNG.Models.Options;
using RoleDDNG.ViewModels.Interfaces;
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

        public MainViewModel()
        {
            IsBusy = true;

            ShowDiceRollWindow = new RelayCommand(ShowDiceRollWindowMethod);
            ShowTownGeneratorWindow = new RelayCommand(ShowTownGeneratorWindowMethod);
            ExitApp = new AsyncCommand(ExitAppMethodAsync);
            LoadAppSettings = new AsyncCommand(LoadAppSettingsMethodAsync);
            BackgroundSource = SimpleIoc.Default.GetInstance<IBackgroundSource>().GetBackgroundSource();
            LoadAppSettings.Execute(this);
            IsBusy = false;
        }

        public IAsyncCommand LoadAppSettings { get; private set; }

        private async Task LoadAppSettingsMethodAsync()
        {
            if (File.Exists(_appSettingsFilePath))
            {
                var serializer = SimpleIoc.Default.GetInstance<IAsyncSerializer<AppSettings>>();
                AppSettings = await serializer.DeserializeAsync(_appSettingsFilePath).ConfigureAwait(true);
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

        private string _backgroundSource = string.Empty;

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

            await SimpleIoc.Default.GetInstance<IAsyncSerializer<AppSettings>>().SerializeAsync(_appSettingsFilePath, AppSettings).ConfigureAwait(false);
            IsBusy = false;
        }
    }

#pragma warning restore CA1822
}