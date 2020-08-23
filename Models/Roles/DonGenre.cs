﻿using GalaSoft.MvvmLight;
using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [TableName(nameof(DonGenre))]
    [PrimaryKey(nameof(Nom))]
    public class DonGenre : ObservableObject
    {
        private string? _nom;

        public DonGenre()
        {
        }

        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}