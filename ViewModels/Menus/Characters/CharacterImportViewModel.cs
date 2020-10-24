using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight.Ioc;

using System;
using System.IO;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Characters
{
    public class CharacterImportViewModel : CharactersListViewModel, IProgress<Tuple<double, string>>
    {
        private bool _withObjects;
        private string _sourceDbFile;

        public string SourceDbFile { get => Path.GetFileName(_sourceDbFile); set { Set(nameof(SourceDbFile), ref _sourceDbFile, value); } }

        private string _targetDbFile;

        public string TargetDbFile { get => Path.GetFileName(_targetDbFile); set { Set(nameof(TargetDbFile), ref _targetDbFile, value); } }

        public CharacterImportViewModel(string sourceDbFile) : base(sourceDbFile)
        {
            _sourceDbFile = sourceDbFile;
            _targetDbFile = SimpleIoc.Default.GetInstance<MainViewModel>().CurrentCharacterDb;
            DoImport = new AsyncCommand(async () => await DoImportAsync().ConfigureAwait(true));
        }

        private async Task DoImportAsync()
        {
            IsBusy = true;
            await Task.Delay(0).ConfigureAwait(true);
            Report(Tuple.Create(100d, ""));
            IsBusy = false;
        }

        public AsyncCommand DoImport { get; private set; }

        public bool WithObjects { get => _withObjects; set { Set(nameof(WithObjects), ref _withObjects, value); } }

        private double _importPercentage;

        public double ImportPercentage { get => _importPercentage; set { Set(nameof(ImportPercentage), ref _importPercentage, value); } }

        private string _lastImportedEntryName = "";

        public string LastImportedEntryName { get => _lastImportedEntryName; set { Set(nameof(LastImportedEntryName), ref _lastImportedEntryName, value); } }

        public void Report(Tuple<double, string> value)
        {
            if (value is null)
            {
                return;
            }
            _lastImportedEntryName = value.Item2;
            _importPercentage = value.Item1;
        }
    }
}