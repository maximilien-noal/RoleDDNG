﻿using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    public class Resistance : ObservableObject
    {
        private string? _nom;

        public Resistance()
        {
        }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}