using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    public class ArchetypeEquipement : ObservableObject
    {
        private string? _archetype;

        private int _emplacement;

        private int _multiple;

        private string? _objet;

        public ArchetypeEquipement()
        {
        }

        [Column]
        public string? Archetype { get => _archetype; set { Set(nameof(Archetype), ref _archetype, value); } }

        [Column]
        public int Emplacement { get => _emplacement; set { Set(nameof(Emplacement), ref _emplacement, value); } }

        [Column]
        public int Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }

        [Column]
        public string? Objet { get => _objet; set { Set(nameof(Objet), ref _objet, value); } }
    }
}