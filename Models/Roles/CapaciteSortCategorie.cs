namespace RoleDDNG.Models.Roles
{
    public class CapaciteSortCategorie : ObservableObject
    {
        private string? _nom;

        public CapaciteSortCategorie()
        {
        }

        /// <summary> Key </summary>
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}