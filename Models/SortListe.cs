using GalaSoft.MvvmLight;

namespace RoleDDNg.Models
{
    public class SortListe : ObservableObject
    {
        private int _id = 0;
        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }
        private string _nomPerso = "";
        public string NomPerso { get => _nomPerso; set { Set(nameof(NomPerso), ref _nomPerso, value); } }
        private string _classePerso = "";
        public string ClassePerso { get => _classePerso; set { Set(nameof(ClassePerso), ref _classePerso, value); } }
        private string _nomSort = "";
        public string NomSort { get => _nomSort; set { Set(nameof(NomSort), ref _nomSort, value); } }
        private short? _niveau = 1;
        public short? Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }
        private string _metamagie = "";
        public string Metamagie { get => _metamagie; set { Set(nameof(Metamagie), ref _metamagie, value); } }
        private string _emplacementSort = "";
        public string EmplacementSort { get => _emplacementSort; set { Set(nameof(EmplacementSort), ref _emplacementSort, value); } }
    }
}