using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class Alignment : ObservableObject
    {
        private string _nom = "";

        public Alignment()
        {
        }

        /// <summary> Key </summary>
        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}