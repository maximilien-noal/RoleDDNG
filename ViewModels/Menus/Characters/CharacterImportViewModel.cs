using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.Models.Characters;
using RoleDDNG.ViewModels.DB;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Characters
{
    public class CharacterImportViewModel : CharactersListViewModel, IProgress<Tuple<int, string>>
    {
        private bool _withObjects = true;
        private IEnumerable<string> _sourceDbFiles;

        private string _stateMessage = "";

        public string StateMessage { get => _stateMessage; set { Set(nameof(StateMessage), ref _stateMessage, value); } }

        private bool _canImport = true;

        public bool CanImport { get => _canImport; set { Set(nameof(CanImport), ref _canImport, value); } }

        public IEnumerable<string> SourceDbFiles { get => _sourceDbFiles; set { Set(nameof(SourceDbFiles), ref _sourceDbFiles, value); } }

        private string _targetDbFile;

        public string TargetDbFile { get => Path.GetFileName(_targetDbFile); set { Set(nameof(TargetDbFile), ref _targetDbFile, value); } }

        public CharacterImportViewModel(IEnumerable<string> sourceDbFiles) : base("")
        {
            _sourceDbFiles = sourceDbFiles;
            _targetDbFile = SimpleIoc.Default.GetInstance<MainViewModel>().CurrentCharacterDb;
            DoImport = new AsyncCommand(async () => await DoImportAsync().ConfigureAwait(true));
            SelectAll = new RelayCommand(() => { foreach (var item in Collection) { item.IsSelected = true; } });
            SelectNone = new RelayCommand(() => { foreach (var item in Collection) { item.IsSelected = false; } });
        }

        public override async Task LoadDbDataAsync()
        {
            if (SourceDbFiles.Any() == false)
            {
                return;
            }
            try
            {
                IsBusy = true;
                foreach (var source in SourceDbFiles)
                {
                    await base.LoadDbDataFromFileAsync(source).ConfigureAwait(true);
                }
                IsBusy = false;
            }
            catch
            {
                Cancel.Execute(this);
                throw;
            }
        }

        public async Task SetImportNamesAsync()
        {
            using var targetDb = DB.CharactersDb.Create();
            IsBusy = true;
            var existingCharacters = new List<Personnage>();
            using var elementsReader = await targetDb.QueryAsync<Personnage>(DB.CommonQueries.DbCharactersNames).ConfigureAwait(true);
            while (await elementsReader.ReadAsync().ConfigureAwait(true))
            {
                existingCharacters.Add(elementsReader.Poco);
            }
            IsBusy = false;
            for (int i = 0; i < Collection.Count; i++)
            {
                var character = Collection[i];
                var fallbackName = string.IsNullOrWhiteSpace(character.Nom) ? $"Sans nom {i}" : character.Nom;
                var num = 2;
                while (existingCharacters.Any(x => x.Nom == fallbackName))
                {
                    fallbackName = $"{fallbackName} {num++}";
                }
                character.ImportName = fallbackName;
            }
        }

        public RelayCommand SelectAll { get; private set; }
        public RelayCommand SelectNone { get; private set; }

        private async Task ImportObjectsAsync()
        {
            if (WithObjects == false)
            {
                return;
            }
            await Task.Delay(0).ConfigureAwait(true);
        }

        private async Task DoImportAsync()
        {
            CanImport = false;
            var objetImportTask = ImportObjectsAsync();
            var updateImportNamesTask = SetImportNamesAsync();
            await Task.WhenAll(objetImportTask, updateImportNamesTask).ConfigureAwait(true);
            Report(Tuple.Create(0, "Importation en cours..."));
            for (int i = 0; i < Collection.Count; i++)
            {
                var character = Collection[i];
                var percentage = (i + 1) / Collection.Count * 100;
                Report(Tuple.Create(percentage, $"{character.Nom} importé sous le nom {character.ImportName}."));
            }
            CanImport = true;
        }

        public AsyncCommand DoImport { get; private set; }

        public bool WithObjects { get => _withObjects; set { Set(nameof(WithObjects), ref _withObjects, value); } }

        private int _percentage;

        public int Percentage { get => _percentage; set { Set(nameof(Percentage), ref _percentage, value); } }

        public void Report(Tuple<int, string> value)
        {
            if (value is null)
            {
                return;
            }
            Percentage = value.Item1;
            StateMessage = value.Item2;
        }
    }
}