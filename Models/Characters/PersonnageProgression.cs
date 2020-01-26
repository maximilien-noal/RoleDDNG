using System.ComponentModel.DataAnnotations.Schema;

using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Characters
{
    public class PersonnageProgression : ObservableObject
    {
        private short? _cha = 0;

        private string _classe = "";

        private string _competence = "";

        private short? _con = 0;

        private short? _dex = 0;

        private short? _for = 0;

        private int _id = 0;

        private short? _int = 0;

        private short? _modifint = 0;

        private short? _niv = 0;

        private short? _niveau = 0;

        private string _nom = "";

        private Personnage _nomNavigation = new Personnage();

        private short? _pv = 0;

        private short? _sag = 0;

        public short? Cha { get => _cha; set { Set(nameof(Cha), ref _cha, value); } }

        public string Classe { get => _classe; set { Set(nameof(Classe), ref _classe, value); } }

        public string Competence { get => _competence; set { Set(nameof(Competence), ref _competence, value); } }

        public short? Con { get => _con; set { Set(nameof(Con), ref _con, value); } }

        public short? Dex { get => _dex; set { Set(nameof(Dex), ref _dex, value); } }

        public short? For { get => _for; set { Set(nameof(For), ref _for, value); } }

        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }

        [Column("Int")]
        public short? Intelligence { get => _int; set { Set(nameof(Intelligence), ref _int, value); } }

        public short? Modifint { get => _modifint; set { Set(nameof(Modifint), ref _modifint, value); } }

        public short? Niv { get => _niv; set { Set(nameof(Niv), ref _niv, value); } }

        public short? Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }

        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        public virtual Personnage NomNavigation { get => _nomNavigation; set { Set(nameof(NomNavigation), ref _nomNavigation, value); } }

        public short? Pv { get => _pv; set { Set(nameof(Pv), ref _pv, value); } }

        public short? Sag { get => _sag; set { Set(nameof(Sag), ref _sag, value); } }
    }
}