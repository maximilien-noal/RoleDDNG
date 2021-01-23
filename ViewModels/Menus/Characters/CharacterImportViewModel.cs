using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using PetaPoco;

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

#pragma warning disable CA1822 // Mark members as static (used for binding)
        public string TargetDbFile => Path.GetFileName(SimpleIoc.Default.GetInstance<MainViewModel>().CurrentCharacterDb);
#pragma warning restore CA1822 // Mark members as static (used for binding)

        public CharacterImportViewModel(IEnumerable<string> sourceDbFiles) : base("")
        {
            _sourceDbFiles = sourceDbFiles;
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

        public async Task SetImportNamesAsync(string targetDbFileName = "")
        {
            if (string.IsNullOrWhiteSpace(targetDbFileName))
            {
                targetDbFileName = SimpleIoc.Default.GetInstance<MainViewModel>().CurrentCharacterDb;
            }
            IsBusy = true;
            var existingCharacters = await DB.DatabaseWrapper.GetCollectionFromQueryAsync<Personnage, List<Personnage>>(CommonQueries.DbCharactersNames, targetDbFileName).ConfigureAwait(true);
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

        private async Task ImportObjectsAsync(string targetDbFileName)
        {
            if (WithObjects == false)
            {
                return;
            }
            Report(Tuple.Create(0, "Importation des objets en cours..."));
            var existingObjects = new Dictionary<string, Objets>();
            var exisitingObjetsRows = (await DB.DatabaseWrapper.GetCollectionFromQueryAsync<Objets, List<Objets>>(DB.CommonQueries.GetAllObjects, targetDbFileName).ConfigureAwait(true))
                .ToDictionary((x) => x.NomObjet ?? "", (x) => x);
            var existingObjectsPropriete = await DB.DatabaseWrapper.GetCollectionFromQueryAsync<ObjetsPropriete, List<ObjetsPropriete>>(DB.CommonQueries.GetAllObjectsPropriete, targetDbFileName).ConfigureAwait(true);
            using var targetDb = DB.DatabaseWrapper.CreateCharactersDb(targetDbFileName);
            var objectsToImport = new Dictionary<string, Objets>();
            var objectsProprieteToImport = new List<ObjetsPropriete>();
            for (int i = 0; i < _sourceDbFiles.Count(); i++)
            {
                var importDbPath = _sourceDbFiles.ElementAt(i);
                var percentage = (i + 1) / Collection.Count * 100;
                Report(Tuple.Create(percentage, $"Objets de la base {Path.GetFileName(importDbPath)} importés"));
                if (File.Exists(importDbPath))
                {
                    var objectsToImportFromDb = (await DB.DatabaseWrapper.GetCollectionFromQueryAsync<Objets, List<Objets>>(DB.CommonQueries.GetAllObjects, importDbPath).ConfigureAwait(true))
                        .ToDictionary((x) => x.NomObjet ?? "", (x) => x);
                    objectsProprieteToImport.AddRange(await DB.DatabaseWrapper.GetCollectionFromQueryAsync<ObjetsPropriete, List<ObjetsPropriete>>(DB.CommonQueries.GetAllObjectsPropriete, importDbPath).ConfigureAwait(true));
                    foreach (var objectToImport in objectsToImport)
                    {
                        if (existingObjects.TryGetValue(objectToImport.Key, out var existingObject) && !existingObject.EqualsTo(objectToImport.Value))
                        {
                            existingObjects[objectToImport.Key] = objectToImport.Value;
                        }
                    }
                }
            }
            await Task.Run(() => targetDb.Execute("DELETE * FROM OBJETS")).ConfigureAwait(true);
            await Task.Run(() => targetDb.Execute("DELETE * FROM OBJETSPROPRIETE")).ConfigureAwait(true);
            await targetDb.BeginTransactionAsync().ConfigureAwait(true);
            foreach (var objectToImport in objectsToImport)
            {
                await Task.Run(() => targetDb.Insert(objectToImport.Value)).ConfigureAwait(true);
            }
            foreach (var objectProprieteToImport in objectsProprieteToImport)
            {
                await Task.Run(() => targetDb.Insert(objectProprieteToImport)).ConfigureAwait(true);
            }
            await Task.Run(() => targetDb.CompleteTransaction()).ConfigureAwait(true);
        }

        private async Task DoImportAsync()
        {
            CanImport = false;
            if (!File.Exists(SimpleIoc.Default.GetInstance<MainViewModel>().CurrentCharacterDb))
            {
                return;
            }
            var currentCharacterDbCopy = Path.GetTempFileName();
            File.Copy(SimpleIoc.Default.GetInstance<MainViewModel>().CurrentCharacterDb, currentCharacterDbCopy, true);
            await ImportObjectsAsync(currentCharacterDbCopy).ConfigureAwait(true);
            await SetImportNamesAsync(currentCharacterDbCopy).ConfigureAwait(true);
            SimpleIoc.Default.GetInstance<ITaskbarProgress>().SetValue(new WindowInteropHelper(Application.Current.MainWindow).Handle, 0, 100);
            Report(Tuple.Create(0, "Importation des personnages en cours..."));
            using var targetDb = DB.DatabaseWrapper.CreateCharactersDb(currentCharacterDbCopy);
            for (int i = 0; i < Collection.Count; i++)
            {
                var character = Collection[i];
                var percentage = (i + 1) / Collection.Count * 100;
                await ImportCharacterAsync(targetDb, character).ConfigureAwait(true);
                Report(Tuple.Create(percentage, $"{character.Nom} importé sous le nom {character.ImportName}."));
            }
            await Task.Run(() => targetDb.CompleteTransaction()).ConfigureAwait(true);
            File.Copy(currentCharacterDbCopy, SimpleIoc.Default.GetInstance<MainViewModel>().CurrentCharacterDb, true);
            CanImport = true;
        }

        private async Task ImportCharacterAsync(Database targetDb, Personnage character)
        {
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