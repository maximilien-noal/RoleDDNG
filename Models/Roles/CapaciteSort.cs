using GalaSoft.MvvmLight;

using System.ComponentModel.DataAnnotations.Schema;

namespace RoleDDNG.Models.Roles
{
    public class CapaciteSort : ObservableObject
    {
        private string _nom = "";

        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        private int _index = 0;

        public int Index { get => _index; set { Set(nameof(Index), ref _index, value); } }

        private int _version = 0;

        public int Version { get => _version; set { Set(nameof(Version), ref _version, value); } }

        private string _genre = "";

        public string Genre { get => _genre; set { Set(nameof(Genre), ref _genre, value); } }

        private string _explication = "";

        public string Explication { get => _explication; set { Set(nameof(Explication), ref _explication, value); } }

        private int _min = 0;

        public int Min { get => _min; set { Set(nameof(Min), ref _min, value); } }

        private int _max = 0;

        public int Max { get => _max; set { Set(nameof(Max), ref _max, value); } }

        private string _bonusBase = "";

        public string BonusBase { get => _bonusBase; set { Set(nameof(BonusBase), ref _bonusBase, value); } }

        private string _augmentation = "";

        public string Augmentation { get => _augmentation; set { Set(nameof(Augmentation), ref _augmentation, value); } }

        private int _tousLes = 0;

        public int TousLes { get => _tousLes; set { Set(nameof(TousLes), ref _tousLes, value); } }

        private string _categorie = "";

        public string Categorie { get => _categorie; set { Set(nameof(Categorie), ref _categorie, value); } }

        private string _sousCategorie = "";

        public string SousCategorie { get => _sousCategorie; set { Set(nameof(SousCategorie), ref _sousCategorie, value); } }

        private string _type = "";

        public string Type { get => _type; set { Set(nameof(Type), ref _type, value); } }

        private bool _lieArme = false;

        [Column("lieArme")]
        public bool LieArme { get => _lieArme; set { Set(nameof(LieArme), ref _lieArme, value); } }

        private bool _multiple = false;

        public bool Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }

        private bool _bonusCapacite = false;

        public bool BonusCapacite { get => _bonusCapacite; set { Set(nameof(BonusCapacite), ref _bonusCapacite, value); } }

        private bool _personnel = false;

        public bool Personnel { get => _personnel; set { Set(nameof(Personnel), ref _personnel, value); } }
    }
}