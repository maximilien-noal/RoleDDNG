using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class ClasseCorrespondanceMagie : ObservableObject
    {
        private int _bonusDeSort;

        private string? _classeEsclave;

        private string? _classeMaitre;

        private int _niveau;

        private int _point;

        private int _sortAutreEnPlus;

        public ClasseCorrespondanceMagie()
        {
        }

        public int BonusDeSort { get => _bonusDeSort; set { Set(nameof(BonusDeSort), ref _bonusDeSort, value); } }

        public string? ClasseEsclave { get => _classeEsclave; set { Set(nameof(ClasseEsclave), ref _classeEsclave, value); } }

        public string? ClasseMaitre { get => _classeMaitre; set { Set(nameof(ClasseMaitre), ref _classeMaitre, value); } }

        public int Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }

        public int Point { get => _point; set { Set(nameof(Point), ref _point, value); } }

        public int SortAutreEnPlus { get => _sortAutreEnPlus; set { Set(nameof(SortAutreEnPlus), ref _sortAutreEnPlus, value); } }
    }
}