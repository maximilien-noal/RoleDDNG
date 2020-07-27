using PetaPoco;

using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class Archetype : ObservableObject
    {
        private short? _adjNiv = 0;

        private string _archetype = "";

        private int _ca = 0;

        private int _cha = 0;

        private int _competence = 0;

        private int _con = 0;

        private int _deplacement = 0;

        private int _deplacementAir = 0;

        private int _deplacementEau = 0;

        private string _description = "";

        private int _dex = 0;

        private int _don = 0;

        private string _don1 = "";

        private string _don10 = "";

        private string _don11 = "";

        private string _don12 = "";

        private string _don2 = "";

        private string _don3 = "";

        private string _don4 = "";

        private string _don5 = "";

        private string _don6 = "";

        private string _don7 = "";

        private string _don8 = "";

        private string _don9 = "";

        private int _for = 0;

        private int _fp = 0;

        private int _int = 0;

        private string _langue = "";

        private string _manoeuvrabilite = "";

        private int _number = 0;

        private int _progressionResistance = 0;

        private int _reflexe = 0;

        private int _resistanceAcide = 0;

        private int _resistanceDegat = 0;

        private string _resistanceDegatTexte = "";

        private int _resistanceElectricite = 0;

        private int _resistanceFeu = 0;

        private int _resistanceFroid = 0;

        private int _resistanceMagie = 0;

        private int _resistanceSon = 0;

        private int _sag = 0;

        private string _taille = "";

        private int _vigueur = 0;

        private int _vision = 0;

        private string _vision_texte = "";

        private int _volonte = 0;

        public Archetype()
        {
        }

        public short? AdjNiv { get => _adjNiv; set { Set(nameof(AdjNiv), ref _adjNiv, value); } }

        public int Ca { get => _ca; set { Set(nameof(Ca), ref _ca, value); } }

        public int Cha { get => _cha; set { Set(nameof(Cha), ref _cha, value); } }

        public int Competence { get => _competence; set { Set(nameof(Competence), ref _competence, value); } }

        public int Con { get => _con; set { Set(nameof(Con), ref _con, value); } }

        [Column("Deplacement_Eau")]
        public int DeplacemeentEau { get => _deplacementEau; set { Set(nameof(DeplacemeentEau), ref _deplacementEau, value); } }

        public int Deplacement { get => _deplacement; set { Set(nameof(Deplacement), ref _deplacement, value); } }

        [Column("Deplacement_Air")]
        public int DeplacementAir { get => _deplacementAir; set { Set(nameof(DeplacementAir), ref _deplacementAir, value); } }

        public string Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        public int Dex { get => _dex; set { Set(nameof(Dex), ref _dex, value); } }

        public int Don { get => _don; set { Set(nameof(Don), ref _don, value); } }

        [Column("Don_1")]
        public string Don1 { get => _don1; set { Set(nameof(Don1), ref _don1, value); } }

        [Column("Don_10")]
        public string Don10 { get => _don10; set { Set(nameof(Don10), ref _don10, value); } }

        [Column("Don_11")]
        public string Don11 { get => _don11; set { Set(nameof(Don11), ref _don11, value); } }

        [Column("Don_12")]
        public string Don12 { get => _don12; set { Set(nameof(Don12), ref _don12, value); } }

        [Column("Don_2")]
        public string Don2 { get => _don2; set { Set(nameof(Don2), ref _don2, value); } }

        [Column("Don_3")]
        public string Don3 { get => _don3; set { Set(nameof(Don3), ref _don3, value); } }

        [Column("Don_4")]
        public string Don4 { get => _don4; set { Set(nameof(Don4), ref _don4, value); } }

        [Column("Don_5")]
        public string Don5 { get => _don5; set { Set(nameof(Don5), ref _don5, value); } }

        [Column("Don_6")]
        public string Don6 { get => _don6; set { Set(nameof(Don6), ref _don6, value); } }

        [Column("Don_7")]
        public string Don7 { get => _don7; set { Set(nameof(Don7), ref _don7, value); } }

        [Column("Don_8")]
        public string Don8 { get => _don8; set { Set(nameof(Don8), ref _don8, value); } }

        [Column("Don_9")]
        public string Don9 { get => _don9; set { Set(nameof(Don9), ref _don9, value); } }

        public int For { get => _for; set { Set(nameof(For), ref _for, value); } }

        public int FP { get => _fp; set { Set(nameof(FP), ref _fp, value); } }

        [Column("Int")]
        public int Intelligence { get => _int; set { Set(nameof(Intelligence), ref _int, value); } }

        public string Langue { get => _langue; set { Set(nameof(Langue), ref _langue, value); } }

        public string Manoeuvrabilite { get => _manoeuvrabilite; set { Set(nameof(Manoeuvrabilite), ref _manoeuvrabilite, value); } }

        /// <summary> Key </summary>
        [Column("Archetype")]
        public string NomArchetype { get => _archetype; set { Set(nameof(NomArchetype), ref _archetype, value); } }

        public int Number { get => _number; set { Set(nameof(Number), ref _number, value); } }

        [Column("Progression_rm")]
        public int ProgressionResistance { get => _progressionResistance; set { Set(nameof(ProgressionResistance), ref _progressionResistance, value); } }

        public int Reflexe { get => _reflexe; set { Set(nameof(Reflexe), ref _reflexe, value); } }

        [Column("Resistance_acide")]
        public int ResistanceAcide { get => _resistanceAcide; set { Set(nameof(ResistanceAcide), ref _resistanceAcide, value); } }

        [Column("Resistance_degat")]
        public int ResistanceDegat { get => _resistanceDegat; set { Set(nameof(ResistanceDegat), ref _resistanceDegat, value); } }

        [Column("Rd_texte")]
        public string ResistanceDegatTexte { get => _resistanceDegatTexte; set { Set(nameof(ResistanceDegatTexte), ref _resistanceDegatTexte, value); } }

        [Column("Resistance_electricite")]
        public int ResistanceElectricite { get => _resistanceElectricite; set { Set(nameof(ResistanceElectricite), ref _resistanceElectricite, value); } }

        [Column("Resistance_feu")]
        public int ResistanceFeu { get => _resistanceFeu; set { Set(nameof(ResistanceFeu), ref _resistanceFeu, value); } }

        [Column("Resistance_froid")]
        public int ResistanceFroid { get => _resistanceFroid; set { Set(nameof(ResistanceFroid), ref _resistanceFroid, value); } }

        [Column("Resistance_magie")]
        public int ResistanceMagie { get => _resistanceMagie; set { Set(nameof(ResistanceMagie), ref _resistanceMagie, value); } }

        [Column("Resistance_son")]
        public int ResistanceSon { get => _resistanceSon; set { Set(nameof(ResistanceSon), ref _resistanceSon, value); } }

        public int Sag { get => _sag; set { Set(nameof(Sag), ref _sag, value); } }

        public string Taille { get => _taille; set { Set(nameof(Taille), ref _taille, value); } }

        public int Vigueur { get => _vigueur; set { Set(nameof(Vigueur), ref _vigueur, value); } }

        public int Vision { get => _vision; set { Set(nameof(Vision), ref _vision, value); } }

        [Column("Vision_Texte")]
        public string VisionTexte { get => _vision_texte; set { Set(nameof(VisionTexte), ref _vision_texte, value); } }

        public int Volonte { get => _volonte; set { Set(nameof(Volonte), ref _volonte, value); } }
    }
}