using GalaSoft.MvvmLight;
using System;
using System.Collections.Generic;
using System.Text;

namespace RoleDDNG.Models.Roles
{
    public class Alignements : ObservableObject
    {
        private string _nom = "";

        /// <summary>
        /// Key
        /// </summary>
        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        private int _ordre = 0;

        public int Ordre { get => _ordre; set { Set(nameof(Ordre), ref _ordre, value); } }

        private string _description = "";

        public string Description { get => _description; set { Set(nameof(Description), ref _description, value); } }
    }
}