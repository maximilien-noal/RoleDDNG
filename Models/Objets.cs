using GalaSoft.MvvmLight;

namespace RoleDDNg.Models
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
        public double? EchecSortProfane { get; set; }
        public short? PenaliteArmure { get; set; }
        public short? BonusDexMax { get; set; }
        public string GroupeObjet { get; set; }
        public string Taille { get; set; }
        public string Type { get; set; }
        public string DegatDes { get; set; }
        public short? ZoneCritique { get; set; }
        public short? MultCrit { get; set; }
        public double? Portee { get; set; }
        public string Description { get; set; }
        public short? Solidite { get; set; }
        public short? Resistance { get; set; }
        public double? CoutTotal { get; set; }
        public short? Charge { get; set; }
    }
}