using PetaPoco;

namespace RoleDDNG.Models.Characters
{
    public class PersonnageProgression : ObservableObject
    {
        private short? _cha;

        private string? _classe;

        private string? _competence;

        private short? _con;

        private short? _dex;

        private short? _for;

        private int _id;

        private short? _int;

        private short? _modifint;

        private short? _niv;

        private short? _niveau;

        private string? _nom;

        private short? _pv;

        private short? _sag;

        public PersonnageProgression()
        {
        }

        [Column]
        public short? Cha { get => _cha; set { Set(nameof(Cha), ref _cha, value); } }

        [Column]
        public string? Classe { get => _classe; set { Set(nameof(Classe), ref _classe, value); } }

        [Column]
        public string? Competence { get => _competence; set { Set(nameof(Competence), ref _competence, value); } }

        [Column]
        public short? Con { get => _con; set { Set(nameof(Con), ref _con, value); } }

        [Column]
        public short? Dex { get => _dex; set { Set(nameof(Dex), ref _dex, value); } }

        [Column]
        public short? For { get => _for; set { Set(nameof(For), ref _for, value); } }

        [Column]
        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }

        [Column("Int")]
        public short? Intelligence { get => _int; set { Set(nameof(Intelligence), ref _int, value); } }

        [Column]
        public short? Modifint { get => _modifint; set { Set(nameof(Modifint), ref _modifint, value); } }

        [Column]
        public short? Niv { get => _niv; set { Set(nameof(Niv), ref _niv, value); } }

        [Column]
        public short? Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        [Column]
        public short? Pv { get => _pv; set { Set(nameof(Pv), ref _pv, value); } }

        [Column]
        public short? Sag { get => _sag; set { Set(nameof(Sag), ref _sag, value); } }
    }
}