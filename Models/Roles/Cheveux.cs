using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [PrimaryKey(nameof(Nom))]
    public class Cheveux : ObservableObject
    {
        private string? _nom;

        public Cheveux()
        {
        }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}