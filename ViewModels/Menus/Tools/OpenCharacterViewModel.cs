﻿using GalaSoft.MvvmLight;
using RoleDDNG.Models.Characters;
using RoleDDNG.ViewModels.Interfaces;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Menus.Tools
{
    public class OpenCharacterViewModel : ViewModelBase, IDocumentViewModel, IDbDependantViewModel
    {
        private const string DbCharactersQuery = "select nom,image,race,niv_1,niv_2,niv_3,niv_4,niv_5,niv_6,niv_7,niv_8,classe_1,classe_2,classe_3,classe_4,classe_5,classe_6,classe_7,classe_8 from personnage where exclu=false order by nom";

        private ObservableCollection<Personnage> _characters = new ObservableCollection<Personnage>();

        private bool _isBusy = false;

        public ObservableCollection<Personnage> Characters { get => _characters; private set { Set(nameof(Characters), ref _characters, value); } }

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public async Task LoadDbDataAsync()
        {
            IsBusy = true;
            using var charactersDb = DB.CharactersDb.Create();
            Characters = new ObservableCollection<Personnage>(await Task.Run(() => charactersDb.Query<Personnage>(DbCharactersQuery)).ConfigureAwait(false));
            IsBusy = false;
        }
    }
}