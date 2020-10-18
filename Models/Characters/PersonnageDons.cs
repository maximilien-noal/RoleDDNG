using PetaPoco;

namespace RoleDDNG.Models.Characters
{
    [TableName(nameof(PersonnageDons))]
    public class PersonnageDons : ObservableObject
    {
        private string? _classe;

        private string? _dons;

        private string? _libelle;

        private short? _niv = 0;

        private short? _niveau = 0;

        private string? _nom;

        public PersonnageDons()
        {
        }

        [Column("classe")]
        public string? Classe { get => _classe; set { Set(nameof(Classe), ref _classe, value); } }

        [Column]
        public string? Dons { get => _dons; set { Set(nameof(Dons), ref _dons, value); } }

        [Column]
        public string? Libelle { get => _libelle; set { Set(nameof(Libelle), ref _libelle, value); } }

        [Column("niv")]
        public short? Niv { get => _niv; set { Set(nameof(Niv), ref _niv, value); } }

        [Column]
        public short? Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }

        [Column("nom")]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}