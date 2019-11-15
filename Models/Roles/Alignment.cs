using GalaSoft.MvvmLight;
using System;
using System.Collections.Generic;
using System.Text;

namespace RoleDDNG.Models.Roles
{
    public class Alignment : ObservableObject
    {
        private string _nom = "";

        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}