using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    public class Yeux : ObservableObject
    {
        private string? _nom;

        public Yeux()
        {
        }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}