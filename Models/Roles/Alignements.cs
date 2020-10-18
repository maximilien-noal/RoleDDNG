using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [PrimaryKey(nameof(Nom))]
    public class Alignements : ObservableObject
    {
        private string? _description;

        private string? _nom;

        private int _ordre;

        public Alignements()
        {
        }

        [Column]
        public string? Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        [Column]
        public int Ordre { get => _ordre; set { Set(nameof(Ordre), ref _ordre, value); } }
    }
}