﻿using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Characters
{
    public class Donsperso : ObservableObject
    {
        private string _caracteristique = "";

        private bool _epique = false;

        private string _explication = "";

        private string _genre = "";

        private bool _liearme = false;

        private bool _multiple = false;

        private string _nom = "";

        private bool _pouvoir = false;

        private bool _prive = false;

        private string _résumé = "";

        private string _source = "";

        private bool _tare = false;

        private bool _trait = false;

        private double? _version = 0d;

        public Donsperso()
        {
        }

        public string Caracteristique { get => _caracteristique; set { Set(nameof(Caracteristique), ref _caracteristique, value); } }

        public bool Epique { get => _epique; set { Set(nameof(Epique), ref _epique, value); } }

        public string Explication { get => _explication; set { Set(nameof(Explication), ref _explication, value); } }

        public string Genre { get => _genre; set { Set(nameof(Genre), ref _genre, value); } }

        public bool Liearme { get => _liearme; set { Set(nameof(Liearme), ref _liearme, value); } }

        public bool Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }

        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        public bool Pouvoir { get => _pouvoir; set { Set(nameof(Pouvoir), ref _pouvoir, value); } }

        public bool Prive { get => _prive; set { Set(nameof(Prive), ref _prive, value); } }

        public string Résumé { get => _résumé; set { Set(nameof(Résumé), ref _résumé, value); } }

        public string Source { get => _source; set { Set(nameof(Source), ref _source, value); } }

        public bool Tare { get => _tare; set { Set(nameof(Tare), ref _tare, value); } }

        public bool Trait { get => _trait; set { Set(nameof(Trait), ref _trait, value); } }

        public double? Version { get => _version; set { Set(nameof(Version), ref _version, value); } }
    }
}