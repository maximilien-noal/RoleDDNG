using GalaSoft.MvvmLight;

namespace RoleDDNG.Models
{
    public class Objets : ObservableObject
    {
        private string _nomObjet = "";
        public string NomObjet { get => _nomObjet; set { Set(nameof(NomObjet), ref _nomObjet, value); } }
        private double? _poidsBase = 0d;
        public double? PoidsBase { get => _poidsBase; set { Set(nameof(PoidsBase), ref _poidsBase, value); } }
        private short? _classArmure = 0;
        public short? ClasseArmure { get => _classArmure; set { Set(nameof(ClasseArmure), ref _classArmure, value); } }
        private string _typeCa = "";
        public string TypeCa { get => _typeCa; set { Set(nameof(TypeCa), ref _typeCa, value); } }
        private double? _echechSortProfane = 0d;
        public double? EchecSortProfane { get => _echechSortProfane; set { Set(nameof(EchecSortProfane), ref _echechSortProfane, value); } }
        private short? _penaliteArmure = 0;
        public short? PenaliteArmure { get => _penaliteArmure; set { Set(nameof(PenaliteArmure), ref _penaliteArmure, value); } }
        private short? _bonusDexMax = 0;
        public short? BonusDexMax { get => _bonusDexMax; set { Set(nameof(BonusDexMax), ref _bonusDexMax, value); } }
        private string _groupObjet = "";
        public string GroupeObjet { get => _groupObjet; set { Set(nameof(GroupeObjet), ref _groupObjet, value); } }
        private string _taille = "";
        public string Taille { get => _taille; set { Set(nameof(Taille), ref _taille, value); } }
        private string _type = "";
        public string Type { get => _type; set { Set(nameof(Type), ref _type, value); } }
        private string _degatDes = "";
        public string DegatDes { get => _degatDes; set { Set(nameof(DegatDes), ref _degatDes, value); } }
        private short? _zoneCritique = 0;
        public short? ZoneCritique { get => _zoneCritique; set { Set(nameof(ZoneCritique), ref _zoneCritique, value); } }
        private short? _multCrit = 0;
        public short? MultCrit { get => _multCrit; set { Set(nameof(MultCrit), ref _multCrit, value); } }
        private double? _portee = 0d;
        public double? Portee { get => _portee; set { Set(nameof(Portee), ref _portee, value); } }
        private string _description = "";
        public string Description { get => _description; set { Set(nameof(Description), ref _description, value); } }
        private short? _solidite = 0;
        public short? Solidite { get => _solidite; set { Set(nameof(Solidite), ref _solidite, value); } }
        private short? _resistance = 0;
        public short? Resistance { get => _resistance; set { Set(nameof(Resistance), ref _resistance, value); } }
        private double? _coutTotal = 0d;
        public double? CoutTotal { get => _coutTotal; set { Set(nameof(CoutTotal), ref _coutTotal, value); } }
        private short? _charge = 0;
        public short? Charge { get => _charge; set { Set(nameof(Charge), ref _charge, value); } }
    }
}