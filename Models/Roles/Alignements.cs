using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class Alignements : ObservableObject
    {
        private string _description = "";

        private string _nom = "";

        private int _ordre;

        public Alignements()
        {
        }

        public string Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        /// <summary> Key </summary>
        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        public int Ordre { get => _ordre; set { Set(nameof(Ordre), ref _ordre, value); } }
    }
}