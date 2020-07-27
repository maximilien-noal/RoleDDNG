using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Characters
{
    public class ObjetsPropriete : ObservableObject
    {
        private int _id = 0;

        private string _nomObjet = "";

        private string _propriete1 = "";

        private string _propriete2 = "";

        private string _propriete3 = "";

        private string _valeur = "";

        public ObjetsPropriete()
        {
        }

        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }

        public string NomObjet { get => _nomObjet; set { Set(nameof(NomObjet), ref _nomObjet, value); } }

        public string Propriete1 { get => _propriete1; set { Set(nameof(Propriete1), ref _propriete1, value); } }

        public string Propriete2 { get => _propriete2; set { Set(nameof(Propriete2), ref _propriete2, value); } }

        public string Propriete3 { get => _propriete3; set { Set(nameof(Propriete3), ref _propriete3, value); } }

        public string Valeur { get => _valeur; set { Set(nameof(Valeur), ref _valeur, value); } }
    }
}