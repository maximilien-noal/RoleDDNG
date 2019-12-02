using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class Caracteristiques : ObservableObject
    {
        private string _nom = "";

        /// <summary>
        /// Key
        /// </summary>
        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        private string _description = "";

        public string Description { get => _description; set { Set(nameof(Description), ref _description, value); } }
    }
}