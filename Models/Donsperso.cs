using GalaSoft.MvvmLight;

namespace RoleDDNG.Models
{
    public class Donsperso : ObservableObject
    {
        private string _nom = "";
        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
        private double? _version = 0d;
        public double? Version { get => _version; set { Set(nameof(Version), ref _version, value); } }
        private string _genre = "";
        public string Genre { get => _genre; set { Set(nameof(Genre), ref _genre, value); } }
        private string _source = "";
        public string Source { get => _source; set { Set(nameof(Source), ref _source, value); } }
        private string _résumé = "";
        public string Résumé { get => _résumé; set { Set(nameof(Résumé), ref _résumé, value); } }
        private string _explication = "";
        public string Explication { get => _explication; set { Set(nameof(Explication), ref _explication, value); } }
        private string _caracteristique = "";
        public string Caracteristique { get => _caracteristique; set { Set(nameof(Caracteristique), ref _caracteristique, value); } }
        private bool _prive = false;
        public bool Prive { get => _prive; set { Set(nameof(Prive), ref _prive, value); } }
        private bool _liearme = false;
        public bool Liearme { get => _liearme; set { Set(nameof(Liearme), ref _liearme, value); } }
        private bool _multiple = false;
        public bool Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }
        private bool _epique = false;
        public bool Epique { get => _epique; set { Set(nameof(Epique), ref _epique, value); } }
        private bool _pouvoir = false;
        public bool Pouvoir { get => _pouvoir; set { Set(nameof(Pouvoir), ref _pouvoir, value); } }
        private bool _tare = false;
        public bool Tare { get => _tare; set { Set(nameof(Tare), ref _tare, value); } }
        private bool _trait = false;
        public bool Trait { get => _trait; set { Set(nameof(Trait), ref _trait, value); } }
    }
}