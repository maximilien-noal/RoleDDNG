using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    public class CapaciteSort : ObservableObject
    {
        private string? _augmentation;

        private string? _bonusBase;

        private bool _bonusCapacite;

        private string? _categorie;

        private string? _explication;

        private string? _genre;

        private int _index;

        private bool _lieArme;

        private int _max;

        private int _min;

        private bool _multiple;

        private string? _nom;

        private bool _personnel;

        private string? _sousCategorie;

        private int _tousLes;

        private string? _type;

        private int _version;

        public CapaciteSort()
        {
        }

        [Column]
        public string? Augmentation { get => _augmentation; set { Set(nameof(Augmentation), ref _augmentation, value); } }

        [Column]
        public string? BonusBase { get => _bonusBase; set { Set(nameof(BonusBase), ref _bonusBase, value); } }

        [Column]
        public bool BonusCapacite { get => _bonusCapacite; set { Set(nameof(BonusCapacite), ref _bonusCapacite, value); } }

        [Column]
        public string? Categorie { get => _categorie; set { Set(nameof(Categorie), ref _categorie, value); } }

        [Column]
        public string? Explication { get => _explication; set { Set(nameof(Explication), ref _explication, value); } }

        [Column]
        public string? Genre { get => _genre; set { Set(nameof(Genre), ref _genre, value); } }

        [Column]
        public int Index { get => _index; set { Set(nameof(Index), ref _index, value); } }

        [Column("lieArme")]
        public bool LieArme { get => _lieArme; set { Set(nameof(LieArme), ref _lieArme, value); } }

        [Column]
        public int Max { get => _max; set { Set(nameof(Max), ref _max, value); } }

        [Column]
        public int Min { get => _min; set { Set(nameof(Min), ref _min, value); } }

        [Column]
        public bool Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        [Column]
        public bool Personnel { get => _personnel; set { Set(nameof(Personnel), ref _personnel, value); } }

        [Column]
        public string? SousCategorie { get => _sousCategorie; set { Set(nameof(SousCategorie), ref _sousCategorie, value); } }

        [Column]
        public int TousLes { get => _tousLes; set { Set(nameof(TousLes), ref _tousLes, value); } }

        [Column]
        public string? Type { get => _type; set { Set(nameof(Type), ref _type, value); } }

        [Column]
        public int Version { get => _version; set { Set(nameof(Version), ref _version, value); } }
    }
}