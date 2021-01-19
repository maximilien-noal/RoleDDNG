using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.Interfaces.Taskbar;
using RoleDDNG.Models.Characters;
using RoleDDNG.ViewModels.DB;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

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
                    await LoadDbDataFromFileAsync(source, CommonQueries.DbCharactersAll).ConfigureAwait(true);
                    foreach (var character in Collection)
                    {
                        character.SourceDb = source;
                    }
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
            IsBusy = true;
            var existingCharacters = await DB.DatabaseWrapper.GetCollectionFromQueryAsync<Personnage, List<Personnage>>(CommonQueries.DbCharactersNames).ConfigureAwait(true);
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
            Report(Tuple.Create(0, "Importation des objets en cours..."));
            var currentObjects = await DB.DatabaseWrapper.GetCollectionFromQueryAsync<Objets, List<Objets>>(DB.CommonQueries.GetObjetsNames).ConfigureAwait(true);
            using var targetDb = DB.DatabaseWrapper.CreateCharactersDb();
            var objectsToImport = new List<Objets>();
            var objectsProprieteToImport = new List<ObjetsPropriete>();
            foreach (var dbPath in _sourceDbFiles)
            {
                if (File.Exists(dbPath))
                {
                    objectsToImport = await DB.DatabaseWrapper.GetCollectionFromQueryAsync<Objets, List<Objets>>(DB.CommonQueries.GetAllObjects, dbPath).ConfigureAwait(true);
                    objectsProprieteToImport = await DB.DatabaseWrapper.GetCollectionFromQueryAsync<ObjetsPropriete, List<ObjetsPropriete>>(DB.CommonQueries.GetAllObjectsPropriete, dbPath).ConfigureAwait(true);
                }
                for (int i = 0; i < objectsToImport.Count; i++)
                {
                    var objectToImport = objectsToImport[i];
                    var matchingObject = currentObjects.FirstOrDefault(x => x.NomObjet == objectToImport.NomObjet);
                    if (!(matchingObject is null) && !matchingObject.EqualsTo(objectToImport))
                    {
                        await Task.Run(() => targetDb.Execute("delete from ObjetsPropriete where [nomObjet]=@0", matchingObject.NomObjet)).ConfigureAwait(true);
                        await Task.Run(() => targetDb.Execute("delete from Objets where [nomObjet]=@0", matchingObject.NomObjet)).ConfigureAwait(true);
                        currentObjects.Remove(matchingObject);
                    }
                    else if (matchingObject is null)
                    {
                        await Task.Run(() => targetDb.Insert(objectToImport)).ConfigureAwait(true);
                    }
                }
                foreach (var objectProprieteToImport in objectsProprieteToImport)
                {
                    await Task.Run(() => targetDb.Insert(objectProprieteToImport)).ConfigureAwait(true);
                }
            }
        }

        private async Task DoImportAsync()
        {
            CanImport = false;
            var objetImportTask = ImportObjectsAsync();
            var updateImportNamesTask = SetImportNamesAsync();
            await Task.WhenAll(objetImportTask, updateImportNamesTask).ConfigureAwait(true);
            SimpleIoc.Default.GetInstance<ITaskbarProgress>().SetValue(new WindowInteropHelper(Application.Current.MainWindow).Handle, 0, 100);
            Report(Tuple.Create(0, "Importation en cours..."));
            for (int i = 0; i < Collection.Count; i++)
            {
                var character = Collection[i];
                var percentage = (i + 1) / Collection.Count * 100;
                await ImportCharacterAsync(character).ConfigureAwait(true);
                Report(Tuple.Create(percentage, $"{character.Nom} importé sous le nom {character.ImportName}."));
            }
            CanImport = true;
        }

        private async Task ImportCharacterAsync(Personnage character)
        {
            var targetDb = DatabaseWrapper.CreateCharactersDb();
            character.Nom = character.ImportName;
            if (WithObjects && File.Exists(character.SourceDb))
            {
                var equipements = await DatabaseWrapper.GetCollectionFromQueryAsync<Equipement, List<Equipement>>("select * from equipement where personnage=@0", character.SourceDb, character.Nom).ConfigureAwait(true);
            }
            await Task.Run(() => targetDb.Insert(character)).ConfigureAwait(true);
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
            SimpleIoc.Default.GetInstance<ITaskbarProgress>().SetValue(new WindowInteropHelper(Application.Current.MainWindow).Handle, Percentage, 100);
            StateMessage = value.Item2;
        }
    }
}