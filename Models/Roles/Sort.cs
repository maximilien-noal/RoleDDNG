using GalaSoft.MvvmLight;
using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [TableName(nameof(Sort))]
    [PrimaryKey(nameof(Nom))]
    public class Sort : ObservableObject
    {
        private string _explication = "";

        private string _nom = "";

        private string _version = "";

        public Sort()
        {
        }

        public string Explication { get => _explication; set { Set(nameof(Explication), ref _explication, value); } }

        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        [Column("version")]
        public string Version { get => _version; set { Set(nameof(Version), ref _version, value); } }
    }
}