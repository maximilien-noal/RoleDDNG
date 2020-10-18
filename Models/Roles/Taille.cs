using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    public class Taille : ObservableObject
    {
        private string? _nom;

        public Taille()
        {
        }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}