using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class CapaciteSortCaracteristique : ObservableObject
    {
        private string _nom = "";

        public CapaciteSortCaracteristique()
        {
        }

        /// <summary> Key </summary>
        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}