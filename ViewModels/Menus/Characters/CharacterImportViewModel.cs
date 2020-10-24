using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.Models.Characters;
using RoleDDNG.ViewModels.DB;
using RoleDDNG.ViewModels.Interfaces;

using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Characters
{
    public class CharacterImportViewModel : ViewModelWithCloseAction<CharacterImportViewModel>, IDocumentViewModel, IDbDependentViewModel, IProgress<Tuple<double, string>>
    {
        private bool _isBusy;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        private bool _withObjects;
        private string _sourceDbFile;

        public string SourceDbFile { get => Path.GetFileName(_sourceDbFile); set { Set(nameof(SourceDbFile), ref _sourceDbFile, value); } }

        private string _targetDbFile;

        public string TargetDbFile { get => Path.GetFileName(_targetDbFile); set { Set(nameof(TargetDbFile), ref _targetDbFile, value); } }

        public CharacterImportViewModel(string sourceDbFile)
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

        private ObservableCollection<Personnage> _characters = new ObservableCollection<Personnage>();

        public ObservableCollection<Personnage> Characters { get => _characters; private set { Set(nameof(Characters), ref _characters, value); } }

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

        public async Task LoadDbDataAsync()
        {
            try
            {
                IsBusy = true;
                Characters.Clear();
                using var charactersDb = CharactersDb.Create(_sourceDbFile);
                using var elementsReader = await charactersDb.QueryAsync<Personnage>(CommonQueries.DbCharactersQuery).ConfigureAwait(true);
                while(await elementsReader.ReadAsync().ConfigureAwait(true))
                {
                    Characters.Add(elementsReader.Poco);
                }
                IsBusy = false;
            }
            catch
            {
                Cancel.Execute(this);
                throw;
            }
        }
    }
}