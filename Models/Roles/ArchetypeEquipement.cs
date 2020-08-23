using GalaSoft.MvvmLight;

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

        public string? Archetype { get => _archetype; set { Set(nameof(Archetype), ref _archetype, value); } }

        public int Emplacement { get => _emplacement; set { Set(nameof(Emplacement), ref _emplacement, value); } }

        public int Multiple { get => _multiple; set { Set(nameof(Multiple), ref _multiple, value); } }

        public string? Objet { get => _objet; set { Set(nameof(Objet), ref _objet, value); } }
    }
}