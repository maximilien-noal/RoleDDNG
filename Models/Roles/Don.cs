using GalaSoft.MvvmLight;
using PetaPoco;
using System;

namespace RoleDDNG.Models.Roles
{
    [TableName("Dons")]
    public class Don : ObservableObject
    {
        private string? _caracteristique;

        private string? _categorie;

        private bool? _competence;

        private string? _condition;

        private DateTime? _date;

        private bool? _epique;

        private string? _explication;

        private string? _genre;

        private string? _js;

        private bool? _lieArme;

        private bool? _multiple;

        private string? _nom;

        private string? _normal;

        private bool? _prive;

        private string? _resume;

        private string? _source;

        private string? _special;

        private bool? _supplementaire;

        private bool? _tare;

        private bool? _trait;

        private string? _type;

        private double? _version;

        public Don()
        {
        }

        public string? Caracteristique { get => _caracteristique; set { Set(nameof(Caracteristique), ref _caracteristique, value); } }

        public string? Categorie { get => _categorie; set { Set(nameof(Categorie), ref _categorie, value); } }

        public bool? Competence { get => _competence; set { Set(nameof(Competence), ref _competence, value); } }

        public string? Condition { get => _condition; set { Set(nameof(Condition), ref _condition, value); } }

        public DateTime? Date { get => _date; set { Set(nameof(Date), ref _date, value); } }

        public bool? Epique { get => _epique; set { Set(nameof(Epique), ref _epique, value); } }

        public string? Explication { get => _explication; set { Set(nameof(Explication), ref _explication, value); } }

        public string? Genre { get => _genre; set { Set(nameof(Genre), ref _genre, value); } }

        public string? JS { get => _js; set { Set(nameof(JS), ref _js, value); } }

        [Column("lieArme")]
        public bool? LieArme { get => _lieArme; set { Set(nameof(LieArme), ref _lieArme, value); } }

        public bool? Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }

        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        public string? Normal { get => _normal; set { Set(nameof(Normal), ref _normal, value); } }

        [Column("prive")]
        public bool? Prive { get => _prive; set { Set(nameof(Prive), ref _prive, value); } }

        [Column("Résumé")]
        public string? Resume { get => _resume; set { Set(nameof(Resume), ref _resume, value); } }

        public string? Source { get => _source; set { Set(nameof(Source), ref _source, value); } }

        [Column("Spécial")]
        public string? Special { get => _special; set { Set(nameof(Special), ref _special, value); } }

        public bool? Supplementaire { get => _supplementaire; set { Set(nameof(Supplementaire), ref _supplementaire, value); } }

        public bool? Tare { get => _tare; set { Set(nameof(Tare), ref _tare, value); } }

        public bool? Trait { get => _trait; set { Set(nameof(Trait), ref _trait, value); } }

        public string? Type { get => _type; set { Set(nameof(Type), ref _type, value); } }

        public double? Version { get => _version; set { Set(nameof(Version), ref _version, value); } }
    }
}