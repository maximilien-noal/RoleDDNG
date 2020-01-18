﻿using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class ClasseCorrespondanceMagie : ObservableObject
    {
        private int _bonusDeSort = 0;

        private string _classeEsclave = "";

        private string _classeMaitre = "";

        private int _niveau = 0;

        private int _point = 0;

        private int _sortAutreEnPlus = 0;

        public int BonusDeSort { get => _bonusDeSort; set { Set(nameof(BonusDeSort), ref _bonusDeSort, value); } }

        public string ClasseEsclave { get => _classeEsclave; set { Set(nameof(ClasseEsclave), ref _classeEsclave, value); } }

        public string ClasseMaitre { get => _classeMaitre; set { Set(nameof(ClasseMaitre), ref _classeMaitre, value); } }

        public int Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }

        public int Point { get => _point; set { Set(nameof(Point), ref _point, value); } }

        public int SortAutreEnPlus { get => _sortAutreEnPlus; set { Set(nameof(SortAutreEnPlus), ref _sortAutreEnPlus, value); } }
    }
}