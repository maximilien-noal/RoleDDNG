using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    public class Terrain : ObservableObject
    {
        private string? _description;

        private string? _nom;

        private string? _type;

        public Terrain()
        {
        }

        [Column]
        public string? Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        [Column]
        public string? Type { get => _type; set { Set(nameof(Type), ref _type, value); } }
    }
}