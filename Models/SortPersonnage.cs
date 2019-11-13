using GalaSoft.MvvmLight;

namespace RoleDDNG.Models
{
    public class SortPersonnage : ObservableObject
    {
        private int _id = 0;
        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }
        private string _nomPerso = "";
        public string NomPerso { get => _nomPerso; set { Set(nameof(NomPerso), ref _nomPerso, value); } }
        private string _classePerso = "";
        public string ClassePerso { get => _classePerso; set { Set(nameof(ClassePerso), ref _classePerso, value); } }
        private string _nomSort = "";
        public string NomSort { get => _nomSort; set { Set(nameof(NomSort), ref _nomSort, value); } }
        private string _niveau = "";
        public string Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }

        private Personnage _nomPersoNavigation = new Personnage();
        public virtual Personnage NomPersoNavigation { get => _nomPersoNavigation; set { Set(nameof(NomPersoNavigation), ref _nomPersoNavigation, value); } }
    }
}