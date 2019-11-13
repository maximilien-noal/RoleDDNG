using GalaSoft.MvvmLight;

namespace RoleDDNG.Models
{
    public class ObjetsPropriete : ObservableObject
    {
        private int _id = 0;
        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }
        private string _nomObjet = "";
        public string NomObjet { get => _nomObjet; set { Set(nameof(NomObjet), ref _nomObjet, value); } }
        private string _propriete1 = "";
        public string Propriete1 { get => _propriete1; set { Set(nameof(Propriete1), ref _propriete1, value); } }
        private string _propriete2 = "";
        public string Propriete2 { get => _propriete2; set { Set(nameof(Propriete2), ref _propriete2, value); } }
        private string _propriete3 = "";
        public string Propriete3 { get => _propriete3; set { Set(nameof(Propriete3), ref _propriete3, value); } }
        private string _valeur = "";
        public string Valeur { get => _valeur; set { Set(nameof(Valeur), ref _valeur, value); } }
    }
}