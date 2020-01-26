using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class Caracteristiques : ObservableObject
    {
        private string _description = "";

        private string _nom = "";

        public string Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        /// <summary>
        /// Key
        /// </summary>
        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}