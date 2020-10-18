namespace RoleDDNG.Models.Roles
{
    public class Cheveux : ObservableObject
    {
        private string? _nom;

        public Cheveux()
        {
        }

        /// <summary> Key </summary>
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}