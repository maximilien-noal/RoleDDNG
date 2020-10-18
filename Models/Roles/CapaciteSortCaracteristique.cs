using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [PrimaryKey(nameof(Nom))]
    public class CapaciteSortCaracteristique : ObservableObject
    {
        private string? _nom;

        public CapaciteSortCaracteristique()
        {
        }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}