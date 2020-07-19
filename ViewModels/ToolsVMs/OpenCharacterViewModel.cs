﻿using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Ioc;
using RoleDDNG.DatabaseLayer;
using RoleDDNG.DatabaseLayer.Enums;
using RoleDDNG.DatabaseLayer.Models;
using RoleDDNG.Models.Characters;
using RoleDDNG.ViewModels.Interfaces;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.ToolsVMs
{
    public class OpenCharacterViewModel : ViewModelBase, IDocumentViewModel, ICharactersDbDependentViewModel
    {
        private const string DbCharactersQuery = "select nom,image,race,niv_1,niv_2,niv_3,niv_4,niv_5,niv_6,niv_7,niv_8,classe_1,classe_2,classe_3,classe_4,classe_5,classe_6,classe_7,classe_8 from personnage where exclu=false order by nom";

        private ObservableCollection<Personnage> _characters = new ObservableCollection<Personnage>();

        private bool _isBusy = false;

        public ObservableCollection<Personnage> Characters { get => _characters; private set { Set(nameof(Characters), ref _characters, value); } }

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public async Task LoadCharactersDbDataAsync()
        {
            IsBusy = true;
            var db = new DbAccessor(new Database(SimpleIoc.Default.GetInstance<MainViewModel>().AppSettings.LastCharacterDBPath, DbType.UserCharactersDb));
            var charactersFromDb = await db.GetQueryDataAsync<Personnage>(DbCharactersQuery).ConfigureAwait(true);
            Characters = new ObservableCollection<Personnage>(charactersFromDb);
            IsBusy = false;
        }
    }
}