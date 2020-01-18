using System.ComponentModel.DataAnnotations.Schema;

using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class CapaciteSort : ObservableObject
    {
        private string _augmentation = "";

        private string _bonusBase = "";

        private bool _bonusCapacite = false;

        private string _categorie = "";

        private string _explication = "";

        private string _genre = "";

        private int _index = 0;

        private bool _lieArme = false;

        private int _max = 0;

        private int _min = 0;

        private bool _multiple = false;

        private string _nom = "";

        private bool _personnel = false;

        private string _sousCategorie = "";

        private int _tousLes = 0;

        private string _type = "";

        private int _version = 0;

        public string Augmentation { get => _augmentation; set { Set(nameof(Augmentation), ref _augmentation, value); } }

        public string BonusBase { get => _bonusBase; set { Set(nameof(BonusBase), ref _bonusBase, value); } }

        public bool BonusCapacite { get => _bonusCapacite; set { Set(nameof(BonusCapacite), ref _bonusCapacite, value); } }

        public string Categorie { get => _categorie; set { Set(nameof(Categorie), ref _categorie, value); } }

        public string Explication { get => _explication; set { Set(nameof(Explication), ref _explication, value); } }

        public string Genre { get => _genre; set { Set(nameof(Genre), ref _genre, value); } }

        public int Index { get => _index; set { Set(nameof(Index), ref _index, value); } }

        [Column("lieArme")]
        public bool LieArme { get => _lieArme; set { Set(nameof(LieArme), ref _lieArme, value); } }

        public int Max { get => _max; set { Set(nameof(Max), ref _max, value); } }

        public int Min { get => _min; set { Set(nameof(Min), ref _min, value); } }

        public bool Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }

        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        public bool Personnel { get => _personnel; set { Set(nameof(Personnel), ref _personnel, value); } }

        public string SousCategorie { get => _sousCategorie; set { Set(nameof(SousCategorie), ref _sousCategorie, value); } }

        public int TousLes { get => _tousLes; set { Set(nameof(TousLes), ref _tousLes, value); } }

        public string Type { get => _type; set { Set(nameof(Type), ref _type, value); } }

        public int Version { get => _version; set { Set(nameof(Version), ref _version, value); } }
    }
}