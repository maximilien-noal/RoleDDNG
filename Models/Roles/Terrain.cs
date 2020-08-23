using GalaSoft.MvvmLight;

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

        public string? Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        public string? Type { get => _type; set { Set(nameof(Type), ref _type, value); } }
    }
}