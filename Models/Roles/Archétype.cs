using GalaSoft.MvvmLight;

using System.ComponentModel.DataAnnotations.Schema;

namespace RoleDDNG.Models.Roles
{
    public class Archétype : ObservableObject
    {
        private string _archetype = "";

        /// <summary>
        /// Key
        /// </summary>
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

        [Column("Int")]
        public int Intelligence { get => _int; set { Set(nameof(Intelligence), ref _int, value); } }

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

        private string _don1 = "";

        [Column("Don_1")]
        public string Don1 { get => _don1; set { Set(nameof(Don1), ref _don1, value); } }

        private string _don2 = "";

        [Column("Don_2")]
        public string Don2 { get => _don2; set { Set(nameof(Don2), ref _don2, value); } }

        private string _don3 = "";

        [Column("Don_3")]
        public string Don3 { get => _don3; set { Set(nameof(Don3), ref _don3, value); } }

        private string _don4 = "";

        [Column("Don_4")]
        public string Don4 { get => _don4; set { Set(nameof(Don4), ref _don4, value); } }

        private string _don5 = "";

        [Column("Don_5")]
        public string Don5 { get => _don5; set { Set(nameof(Don5), ref _don5, value); } }

        private string _don6 = "";

        [Column("Don_6")]
        public string Don6 { get => _don6; set { Set(nameof(Don6), ref _don6, value); } }

        private string _don7 = "";

        [Column("Don_7")]
        public string Don7 { get => _don7; set { Set(nameof(Don7), ref _don7, value); } }

        private string _don8 = "";

        [Column("Don_8")]
        public string Don8 { get => _don8; set { Set(nameof(Don8), ref _don8, value); } }

        private string _don9 = "";

        [Column("Don_9")]
        public string Don9 { get => _don9; set { Set(nameof(Don9), ref _don9, value); } }

        private string _don10 = "";

        [Column("Don_10")]
        public string Don10 { get => _don10; set { Set(nameof(Don10), ref _don10, value); } }

        private string _don11 = "";

        [Column("Don_11")]
        public string Don11 { get => _don11; set { Set(nameof(Don11), ref _don11, value); } }

        private string _don12 = "";

        [Column("Don_12")]
        public string Don12 { get => _don12; set { Set(nameof(Don12), ref _don12, value); } }

        private string _taille = "";

        public string Taille { get => _taille; set { Set(nameof(Taille), ref _taille, value); } }

        private int _deplacementEau = 0;

        [Column("Deplacement_Eau")]
        public int DeplacemeentEau { get => _deplacementEau; set { Set(nameof(DeplacemeentEau), ref _deplacementEau, value); } }

        private int _deplacementAir = 0;

        [Column("Deplacement_Air")]
        public int DeplacementAir { get => _deplacementAir; set { Set(nameof(DeplacementAir), ref _deplacementAir, value); } }

        private string _manoeuvrabilite = "";

        public string Manoeuvrabilite { get => _manoeuvrabilite; set { Set(nameof(Manoeuvrabilite), ref _manoeuvrabilite, value); } }

        private int _vision = 0;

        public int Vision { get => _vision; set { Set(nameof(Vision), ref _vision, value); } }

        private string _vision_texte = "";

        [Column("Vision_Texte")]
        public string VisionTexte { get => _vision_texte; set { Set(nameof(VisionTexte), ref _vision_texte, value); } }

        private int _resistanceDegat = 0;

        [Column("Resistance_degat")]
        public int ResistanceDegat { get => _resistanceDegat; set { Set(nameof(ResistanceDegat), ref _resistanceDegat, value); } }

        private string _resistanceDegatTexte = "";

        [Column("Rd_texte")]
        public string ResistanceDegatTexte { get => _resistanceDegatTexte; set { Set(nameof(ResistanceDegatTexte), ref _resistanceDegatTexte, value); } }

        private int _resistanceMagie = 0;

        [Column("Resistance_magie")]
        public int ResistanceMagie { get => _resistanceMagie; set { Set(nameof(ResistanceMagie), ref _resistanceMagie, value); } }

        private int _progressionResistance = 0;

        [Column("Progression_rm")]
        public int ProgressionResistance { get => _progressionResistance; set { Set(nameof(ProgressionResistance), ref _progressionResistance, value); } }

        private int _resistanceFeu = 0;

        [Column("Resistance_feu")]
        public int ResistanceFeu { get => _resistanceFeu; set { Set(nameof(ResistanceFeu), ref _resistanceFeu, value); } }

        private int _resistanceFroid = 0;

        [Column("Resistance_froid")]
        public int ResistanceFroid { get => _resistanceFroid; set { Set(nameof(ResistanceFroid), ref _resistanceFroid, value); } }

        private int _resistanceElectricite = 0;

        [Column("Resistance_electricite")]
        public int ResistanceElectricite { get => _resistanceElectricite; set { Set(nameof(ResistanceElectricite), ref _resistanceElectricite, value); } }

        private int _resistanceAcide = 0;

        [Column("Resistance_acide")]
        public int ResistanceAcide { get => _resistanceAcide; set { Set(nameof(ResistanceAcide), ref _resistanceAcide, value); } }

        private int _resistanceSon = 0;

        [Column("Resistance_son")]
        public int ResistanceSon { get => _resistanceSon; set { Set(nameof(ResistanceSon), ref _resistanceSon, value); } }

        private string _langue = "";

        public string Langue { get => _langue; set { Set(nameof(Langue), ref _langue, value); } }

        private int _deplacement = 0;

        public int Deplacement { get => _deplacement; set { Set(nameof(Deplacement), ref _deplacement, value); } }

        private int _number = 0;

        public int Number { get => _number; set { Set(nameof(Number), ref _number, value); } }
    }
}