namespace RoleDDNG.Models.Roles
{
    public class Aberration : ObservableObject
    {
        private string? _nom;

        public Aberration()
        {
        }

        /// <summary> Key </summary>
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}