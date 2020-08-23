using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class Type : ObservableObject
    {
        private string? _nom;

        public Type()
        {
        }

        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}