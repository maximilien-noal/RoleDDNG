using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    public class TypeClasse : ObservableObject
    {
        private string? _nom;

        public TypeClasse()
        {
        }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}