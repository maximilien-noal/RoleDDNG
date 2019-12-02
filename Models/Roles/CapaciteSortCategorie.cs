﻿using GalaSoft.MvvmLight;
using System;
using System.Collections.Generic;
using System.Text;

namespace RoleDDNG.Models.Roles
{
    public class CapaciteSortCategorie : ObservableObject
    {
        private string _nom = "";

        /// <summary>
        /// Key
        /// </summary>
        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}