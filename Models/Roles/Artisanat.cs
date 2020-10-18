namespace RoleDDNG.Models.Roles
{
    public class Artisanat : ObservableObject
    {
        private string? _nom;

        public Artisanat()
        {
        }

        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}