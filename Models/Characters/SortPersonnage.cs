using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Characters
{
    public class SortPersonnage : ObservableObject
    {
        private string _classePerso = "";

        private int _id = 0;

        private string _niveau = "";

        private string _nomPerso = "";

        private Personnage _nomPersoNavigation = new Personnage();

        private string _nomSort = "";

        public SortPersonnage()
        {
        }

        public string ClassePerso { get => _classePerso; set { Set(nameof(ClassePerso), ref _classePerso, value); } }

        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }

        public string Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }

        public string NomPerso { get => _nomPerso; set { Set(nameof(NomPerso), ref _nomPerso, value); } }

        public virtual Personnage NomPersoNavigation { get => _nomPersoNavigation; set { Set(nameof(NomPersoNavigation), ref _nomPersoNavigation, value); } }

        public string NomSort { get => _nomSort; set { Set(nameof(NomSort), ref _nomSort, value); } }
    }
}