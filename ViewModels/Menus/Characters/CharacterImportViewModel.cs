using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.Models.Characters;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Characters
{
    public class CharacterImportViewModel : CharactersListViewModel, IProgress<Tuple<int, string>>
    {
        private bool _withObjects;
        private string _sourceDbFile;

        private string _stateMessage = "";

        public string StateMessage { get => _stateMessage; set { Set(nameof(StateMessage), ref _stateMessage, value); } }

        private bool _canImport = true;

        public bool CanImport { get => _canImport; set { Set(nameof(CanImport), ref _canImport, value); } }

        public string SourceDbFile { get => Path.GetFileName(_sourceDbFile); set { Set(nameof(SourceDbFile), ref _sourceDbFile, value); } }

        private string _targetDbFile;

        public string TargetDbFile { get => Path.GetFileName(_targetDbFile); set { Set(nameof(TargetDbFile), ref _targetDbFile, value); } }

        public CharacterImportViewModel(string sourceDbFile) : base(sourceDbFile)
        {
            _sourceDbFile = sourceDbFile;
            _targetDbFile = SimpleIoc.Default.GetInstance<MainViewModel>().CurrentCharacterDb;
            DoImport = new AsyncCommand(async () => await DoImportAsync().ConfigureAwait(true));
            SelectAll = new RelayCommand(() => { foreach (var item in Collection) { item.IsSelected = true; } });
            SelectNone = new RelayCommand(() => { foreach (var item in Collection) { item.IsSelected = false; } });
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
                var item = Collection[i];
                var fallbackName = string.IsNullOrWhiteSpace(item.Nom) ? $"Sans nom {i}" : item.Nom;
                var num = 2;
                while (existingCharacters.Any(x => x.Nom == fallbackName))
                {
                    fallbackName = $"{fallbackName} {num++}";
                }
                item.NameImport = fallbackName;
            }
        }

        public RelayCommand SelectAll { get; private set; }
        public RelayCommand SelectNone { get; private set; }

        private async Task DoImportAsync()
        {
            CanImport = false;
            await SetImportNamesAsync().ConfigureAwait(true);
            Report(Tuple.Create(0, "Importation en cours..."));
            for (int i = 0; i < Collection.Count; i++)
            {
                var character = Collection[i];
                await Task.Delay(1000).ConfigureAwait(true);
                var percentage = (i + 1) / Collection.Count * 100;
                Report(Tuple.Create(percentage, $"{character.Nom} importé."));
                await Task.Delay(1000).ConfigureAwait(true);
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