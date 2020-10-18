namespace RoleDDNG.Models.Characters
{
    public class Equipement : ObservableObject
    {
        private short? _charge = 0;

        private int _id;

        private short? _multiple = 0;

        private string? _nomObjet;

        private string? _personnage;

        private Personnage _personnageNavigation = new Personnage();

        private short? _place = 0;

        private short? _pos = 0;

        private bool _surPersonnage;

        private string? _type;

        public Equipement()
        {
        }

        public short? Charge { get => _charge; set { Set(nameof(Charge), ref _charge, value); } }

        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }

        public short? Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }

        public string? NomObjet { get => _nomObjet; set { Set(nameof(NomObjet), ref _nomObjet, value); } }

        public string? Personnage { get => _personnage; set { Set(nameof(Personnage), ref _personnage, value); } }

        public virtual Personnage PersonnageNavigation { get => _personnageNavigation; set { Set(nameof(PersonnageNavigation), ref _personnageNavigation, value); } }

        public short? Place { get => _place; set { Set(nameof(Place), ref _place, value); } }

        public short? Pos { get => _pos; set { Set(nameof(Pos), ref _pos, value); } }

        public bool SurPersonnage { get => _surPersonnage; set { Set(nameof(SurPersonnage), ref _surPersonnage, value); } }

        public string? Type { get => _type; set { Set(nameof(Type), ref _type, value); } }
    }
}