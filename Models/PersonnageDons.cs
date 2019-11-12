using GalaSoft.MvvmLight;

namespace RoleDDNg.Models
{
    public class PersonnageDons : ObservableObject
    {
        private int _id = 0;
        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }
        private string _nom = "";
        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
        private string _classe = "";
        public string Classe { get => _classe; set { Set(nameof(Classe), ref _classe, value); } }
        private short? _niveau = 0;
        public short? Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }
        private string _dons = "";
        public string Dons { get => _dons; set { Set(nameof(Dons), ref _dons, value); } }
        private string _libelle = "";
        public string Libelle { get => _libelle; set { Set(nameof(Libelle), ref _libelle, value); } }
        private short? _niv = 0;
        public short? Niv { get => _niv; set { Set(nameof(Niv), ref _niv, value); } }
    }
}