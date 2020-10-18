using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [PrimaryKey(nameof(Nom))]
    public class Aberration : ObservableObject
    {
        private string? _nom;

        public Aberration()
        {
        }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}