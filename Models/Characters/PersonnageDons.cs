using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Characters
{
    public class PersonnageDons : ObservableObject
    {
        private string _classe = "";

        private string _dons = "";

        private int _id = 0;

        private string _libelle = "";

        private short? _niv = 0;

        private short? _niveau = 0;

        private string _nom = "";

        public PersonnageDons()
        {
        }

        public string Classe { get => _classe; set { Set(nameof(Classe), ref _classe, value); } }

        public string Dons { get => _dons; set { Set(nameof(Dons), ref _dons, value); } }

        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }

        public string Libelle { get => _libelle; set { Set(nameof(Libelle), ref _libelle, value); } }

        public short? Niv { get => _niv; set { Set(nameof(Niv), ref _niv, value); } }

        public short? Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }

        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}