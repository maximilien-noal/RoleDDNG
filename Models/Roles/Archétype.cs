using GalaSoft.MvvmLight;
using System;
using System.Collections.Generic;
using System.Text;

namespace RoleDDNG.Models.Roles
{
    public class Archétype : ObservableObject
    {
        private string _archetype = "";

        public string Archetype { get => _archetype; set { Set(nameof(Archetype), ref _archetype, value); } }

        private string _description = "";

        public string Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        private int _for = 0;

        public int For { get => _for; set { Set(nameof(For), ref _for, value); } }

        private int _dex = 0;

        public int Dex { get => _dex; set { Set(nameof(Dex), ref _dex, value); } }

        private int _con = 0;

        public int Con { get => _con; set { Set(nameof(Con), ref _con, value); } }

        private int _int = 0;
#pragma warning disable CA1720 // Justification : Legacy column name compatibility
        public int Int { get => _int; set { Set(nameof(Int), ref _int, value); } }
#pragma warning restore CA1720

        private int _sag = 0;

        public int Sag { get => _sag; set { Set(nameof(Sag), ref _sag, value); } }

        private int _cha = 0;

        public int Cha { get => _cha; set { Set(nameof(Cha), ref _cha, value); } }

        private int _ca = 0;

        public int Ca { get => _ca; set { Set(nameof(Ca), ref _ca, value); } }

        private int _competence = 0;

        public int Competence { get => _competence; set { Set(nameof(Competence), ref _competence, value); } }

        private int _don = 0;

        public int Don { get => _don; set { Set(nameof(Don), ref _don, value); } }

        private int _reflexe = 0;

        public int Reflexe { get => _reflexe; set { Set(nameof(Reflexe), ref _reflexe, value); } }

        private int _vigueur = 0;

        public int Vigueur { get => _vigueur; set { Set(nameof(Vigueur), ref _vigueur, value); } }

        private int _volonte = 0;

        public int Volonte { get => _volonte; set { Set(nameof(Volonte), ref _volonte, value); } }

        private int _adjNiv = 0;

        public int AdjNiv { get => _adjNiv; set { Set(nameof(AdjNiv), ref _adjNiv, value); } }

        private int _fp = 0;

        public int FP { get => _fp; set { Set(nameof(FP), ref _fp, value); } }
#pragma warning disable CA1707 // Justification : Legacy column name compatibility

        private string _don_1 = "";

        public string Don_1 { get => _don_1; set { Set(nameof(_don_1), ref _don_1, value); } }

        private string _don_2 = "";

        public string Don_2 { get => _don_2; set { Set(nameof(_don_2), ref _don_2, value); } }
#pragma warning disable CA1707 // Justification : Legacy column name compatibility
    }
}