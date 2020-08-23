using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class SousType : ObservableObject
    {
        private string? _nom;

        public SousType()
        {
        }

        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}