using PetaPoco;

namespace RoleDDNG.Models.Characters
{
    public class Donsperso : ObservableObject
    {
        private string? _caracteristique;

        private bool _epique;

        private string? _explication;

        private string? _genre;

        private bool _liearme;

        private bool _multiple;

        private string? _nom;

        private bool _pouvoir;

        private bool _prive;

        private string? _résumé;

        private string? _source;

        private bool _tare;

        private bool _trait;

        private double? _version;

        public Donsperso()
        {
        }

        [Column]
        public string? Caracteristique { get => _caracteristique; set { Set(nameof(Caracteristique), ref _caracteristique, value); } }

        [Column]
        public bool Epique { get => _epique; set { Set(nameof(Epique), ref _epique, value); } }

        [Column]
        public string? Explication { get => _explication; set { Set(nameof(Explication), ref _explication, value); } }

        [Column]
        public string? Genre { get => _genre; set { Set(nameof(Genre), ref _genre, value); } }

        [Column]
        public bool Liearme { get => _liearme; set { Set(nameof(Liearme), ref _liearme, value); } }

        [Column]
        public bool Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        [Column]
        public bool Pouvoir { get => _pouvoir; set { Set(nameof(Pouvoir), ref _pouvoir, value); } }

        [Column]
        public bool Prive { get => _prive; set { Set(nameof(Prive), ref _prive, value); } }

        [Column]
        public string? Résumé { get => _résumé; set { Set(nameof(Résumé), ref _résumé, value); } }

        [Column]
        public string? Source { get => _source; set { Set(nameof(Source), ref _source, value); } }

        [Column]
        public bool Tare { get => _tare; set { Set(nameof(Tare), ref _tare, value); } }

        [Column]
        public bool Trait { get => _trait; set { Set(nameof(Trait), ref _trait, value); } }

        [Column]
        public double? Version { get => _version; set { Set(nameof(Version), ref _version, value); } }
    }
}