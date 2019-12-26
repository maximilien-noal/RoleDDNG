using System;
using System.Collections.Generic;
using System.Text;
using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class ClasseCorrespondanceMagie : ObservableObject
    {
        private string _classeEsclave = "";

        public string ClasseEsclave { get => _classeEsclave; set { Set(nameof(ClasseEsclave), ref _classeEsclave, value); } }

        private int _niveau = 0;

        public int Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }

        private string _classeMaitre = "";

        public string ClasseMaitre { get => _classeMaitre; set { Set(nameof(ClasseMaitre), ref _classeMaitre, value); } }

        private int _point = 0;

        public int Point { get => _point; set { Set(nameof(Point), ref _point, value); } }

        private int _bonusDeSort = 0;

        public int BonusDeSort { get => _bonusDeSort; set { Set(nameof(BonusDeSort), ref _bonusDeSort, value); } }

        private int _sortAutreEnPlus = 0;

        public int SortAutreEnPlus { get => _sortAutreEnPlus; set { Set(nameof(SortAutreEnPlus), ref _sortAutreEnPlus, value); } }
    }
}