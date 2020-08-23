using GalaSoft.MvvmLight;
using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [TableName("DonsCategorie")]
    [PrimaryKey(nameof(Nom))]
    public class DonCategorie : ObservableObject
    {
        private string _nom = "";

        public DonCategorie()
        {
        }

        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
    }
}