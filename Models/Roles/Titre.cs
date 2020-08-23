using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class Titre : ObservableObject
    {
        private int? _grade = 0;

        private string? _nom;

        private string? _sexe;

        private string? _vecu;

        public Titre()
        {
        }

        public int? Grade { get => _grade; set { Set(nameof(Grade), ref _grade, value); } }

        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        public string? Sexe { get => _sexe; set { Set(nameof(Sexe), ref _sexe, value); } }

        public string? Vecu { get => _vecu; set { Set(nameof(Vecu), ref _vecu, value); } }
    }
}