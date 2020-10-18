using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [PrimaryKey(nameof(Nom))]
    public class Alignment : ObservableObject
    {
        private string? _nom;

        public Alignment()
        {
        }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}