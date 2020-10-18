using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    public class Source : ObservableObject
    {
        private string? _nom;

        public Source()
        {
        }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}