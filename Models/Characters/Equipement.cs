using PetaPoco;

namespace RoleDDNG.Models.Characters
{
    public class Equipement : ObservableObject
    {
        private short? _charge = 0;

        private int _id;

        private short? _multiple = 0;

        private string? _nomObjet;

        private string? _personnage;

        private short? _place = 0;

        private short? _pos = 0;

        private bool _surPersonnage;

        private string? _type;

        public Equipement()
        {
        }

        [Column]
        public short? Charge { get => _charge; set { Set(nameof(Charge), ref _charge, value); } }

        [Column]
        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }

        [Column]
        public short? Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }

        [Column]
        public string? NomObjet { get => _nomObjet; set { Set(nameof(NomObjet), ref _nomObjet, value); } }

        [Column]
        public string? Personnage { get => _personnage; set { Set(nameof(Personnage), ref _personnage, value); } }

        [Column]
        public short? Place { get => _place; set { Set(nameof(Place), ref _place, value); } }

        [Column]
        public short? Pos { get => _pos; set { Set(nameof(Pos), ref _pos, value); } }

        [Column]
        public bool SurPersonnage { get => _surPersonnage; set { Set(nameof(SurPersonnage), ref _surPersonnage, value); } }

        [Column]
        public string? Type { get => _type; set { Set(nameof(Type), ref _type, value); } }
    }
}