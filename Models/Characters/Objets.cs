namespace RoleDDNG.Models.Characters
{
    public class Objets : ObservableObject
    {
        private short? _bonusDexMax = 0;

        private short? _charge = 0;

        private short? _classArmure = 0;

        private double? _coutTotal = 0d;

        private string? _degatDes;

        private string? _description;

        private double? _echechSortProfane = 0d;

        private string? _groupObjet;

        private short? _multCrit = 0;

        private string? _nomObjet;

        private short? _penaliteArmure = 0;

        private double? _poidsBase = 0d;

        private double? _portee = 0d;

        private short? _resistance = 0;

        private short? _solidite = 0;

        private string? _taille;

        private string? _type;

        private string? _typeCa;

        private short? _zoneCritique = 0;

        public Objets()
        {
        }

        public short? BonusDexMax { get => _bonusDexMax; set { Set(nameof(BonusDexMax), ref _bonusDexMax, value); } }

        public short? Charge { get => _charge; set { Set(nameof(Charge), ref _charge, value); } }

        public short? ClasseArmure { get => _classArmure; set { Set(nameof(ClasseArmure), ref _classArmure, value); } }

        public double? CoutTotal { get => _coutTotal; set { Set(nameof(CoutTotal), ref _coutTotal, value); } }

        public string? DegatDes { get => _degatDes; set { Set(nameof(DegatDes), ref _degatDes, value); } }

        public string? Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        public double? EchecSortProfane { get => _echechSortProfane; set { Set(nameof(EchecSortProfane), ref _echechSortProfane, value); } }

        public string? GroupeObjet { get => _groupObjet; set { Set(nameof(GroupeObjet), ref _groupObjet, value); } }

        public short? MultCrit { get => _multCrit; set { Set(nameof(MultCrit), ref _multCrit, value); } }

        public string? NomObjet { get => _nomObjet; set { Set(nameof(NomObjet), ref _nomObjet, value); } }

        public short? PenaliteArmure { get => _penaliteArmure; set { Set(nameof(PenaliteArmure), ref _penaliteArmure, value); } }

        public double? PoidsBase { get => _poidsBase; set { Set(nameof(PoidsBase), ref _poidsBase, value); } }

        public double? Portee { get => _portee; set { Set(nameof(Portee), ref _portee, value); } }

        public short? Resistance { get => _resistance; set { Set(nameof(Resistance), ref _resistance, value); } }

        public short? Solidite { get => _solidite; set { Set(nameof(Solidite), ref _solidite, value); } }

        public string? Taille { get => _taille; set { Set(nameof(Taille), ref _taille, value); } }

        public string? Type { get => _type; set { Set(nameof(Type), ref _type, value); } }

        public string? TypeCa { get => _typeCa; set { Set(nameof(TypeCa), ref _typeCa, value); } }

        public short? ZoneCritique { get => _zoneCritique; set { Set(nameof(ZoneCritique), ref _zoneCritique, value); } }
    }
}