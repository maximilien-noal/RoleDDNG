using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [PrimaryKey(nameof(Nom))]
    public class Caracteristiques : ObservableObject
    {
        private string? _description;

        private string? _nom;

        public Caracteristiques()
        {
        }

        [Column]
        public string? Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}