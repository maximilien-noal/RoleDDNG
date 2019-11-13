using GalaSoft.MvvmLight;

namespace RoleDDNG.Models
{
    public class PersonnageProgression : ObservableObject
    {
        private int _id = 0;
        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }
        private string _nom = "";
        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
        private string _classe = "";
        public string Classe { get => _classe; set { Set(nameof(Classe), ref _classe, value); } }
        private short? _niveau = 0;
        public short? Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }
        private short? _pv = 0;
        public short? Pv { get => _pv; set { Set(nameof(Pv), ref _pv, value); } }
        private short? _for = 0;
        public short? For { get => _for; set { Set(nameof(For), ref _for, value); } }
        private short? _dex = 0;
        public short? Dex { get => _dex; set { Set(nameof(Dex), ref _dex, value); } }
        private short? _con = 0;
        public short? Con { get => _con; set { Set(nameof(Con), ref _con, value); } }
#pragma warning disable CA1720 // Identifier contains type name (Justification : Legacy column name)
        private short? _int = 0;
        public short? Int { get => _int; set { Set(nameof(Int), ref _int, value); } }
#pragma warning restore CA1720 // Identifier contains type name (Justification : Legacy column name)
        private short? _sag = 0;
        public short? Sag { get => _sag; set { Set(nameof(Sag), ref _sag, value); } }
        private short? _cha = 0;
        public short? Cha { get => _cha; set { Set(nameof(Cha), ref _cha, value); } }
        private short? _niv = 0;
        public short? Niv { get => _niv; set { Set(nameof(Niv), ref _niv, value); } }
        private string _competence = "";
        public string Competence { get => _competence; set { Set(nameof(Competence), ref _competence, value); } }
        private short? _modifint = 0;
        public short? Modifint { get => _modifint; set { Set(nameof(Modifint), ref _modifint, value); } }
        private Personnage _nomNavigation = new Personnage();
        public virtual Personnage NomNavigation { get => _nomNavigation; set { Set(nameof(NomNavigation), ref _nomNavigation, value); } }
    }
}