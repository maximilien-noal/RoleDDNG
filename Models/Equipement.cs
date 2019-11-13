using GalaSoft.MvvmLight;

namespace RoleDDNG.Models
{
    public class Equipement : ObservableObject
    {
        private int _id = 0;
        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }
        private string _nomObjet = "";
        public string NomObjet { get => _nomObjet; set { Set(nameof(NomObjet), ref _nomObjet, value); } }
        private string _personnage = "";
        public string Personnage { get => _personnage; set { Set(nameof(Personnage), ref _personnage, value); } }
        private short? _place = 0;
        public short? Place { get => _place; set { Set(nameof(Place), ref _place, value); } }
        private short? _pos = 0;
        public short? Pos { get => _pos; set { Set(nameof(Pos), ref _pos, value); } }
        private string _type = "";
        public string Type { get => _type; set { Set(nameof(Type), ref _type, value); } }
        private short? _charge = 0;
        public short? Charge { get => _charge; set { Set(nameof(Charge), ref _charge, value); } }
        private bool _surPersonnage = false;
        public bool SurPersonnage { get => _surPersonnage; set { Set(nameof(SurPersonnage), ref _surPersonnage, value); } }
        private short? _multiple = 0;
        public short? Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }

        private Personnage _personnageNavigation = new Personnage();
        public virtual Personnage PersonnageNavigation { get => _personnageNavigation; set { Set(nameof(PersonnageNavigation), ref _personnageNavigation, value); } }
    }
}