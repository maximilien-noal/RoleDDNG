using GalaSoft.MvvmLight;
using PetaPoco;
using System;

namespace RoleDDNG.Models.Roles
{
    [TableName("Dons")]
    public class Don : ObservableObject
    {
        private string? _categorie;

        private DateTime? _date;

        private string? _nom;

        private double? _version;

        public Don()
        {
        }

        public string? Categorie { get => _categorie; set { Set(nameof(Categorie), ref _categorie, value); } }

        public DateTime? Date { get => _date; set { Set(nameof(Date), ref _date, value); } }

        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        public double? Version { get => _version; set { Set(nameof(Version), ref _version, value); } }
    }
}