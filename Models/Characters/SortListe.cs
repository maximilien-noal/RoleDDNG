using PetaPoco;

namespace RoleDDNG.Models.Characters
{
    public class SortListe : ObservableObject
    {
        private string? _classePerso;

        private string? _emplacementSort;

        private int _id;

        private string? _metamagie;

        private short? _niveau = 1;

        private string? _nomPerso;

        private string? _nomSort;

        public SortListe()
        {
        }

        [Column]
        public string? ClassePerso { get => _classePerso; set { Set(nameof(ClassePerso), ref _classePerso, value); } }

        [Column]
        public string? EmplacementSort { get => _emplacementSort; set { Set(nameof(EmplacementSort), ref _emplacementSort, value); } }

        [Column]
        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }

        [Column]
        public string? Metamagie { get => _metamagie; set { Set(nameof(Metamagie), ref _metamagie, value); } }

        [Column]
        public short? Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }

        [Column]
        public string? NomPerso { get => _nomPerso; set { Set(nameof(NomPerso), ref _nomPerso, value); } }

        [Column]
        public string? NomSort { get => _nomSort; set { Set(nameof(NomSort), ref _nomSort, value); } }
    }
}