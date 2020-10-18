using PetaPoco;

using System.Collections.ObjectModel;
using System.Linq;

namespace RoleDDNG.Models.Characters
{
    [TableName(nameof(Personnage))]
    [PrimaryKey("Nom")]
    public class Personnage : ObservableObject
    {
        public int EffectiveLevel()
        {
            var col = new short?[] { Niv1, Niv2, Niv3, Niv4, Niv5, Niv6, Niv7, Niv8 };
            var maxLevel = col.Max();
            if (maxLevel is null)
            {
                return 0;
            }
            else
            {
                return (int)maxLevel;
            }
        }

        public string EffectiveClass()
        {
            var levels = new short?[] { Niv1, Niv2, Niv3, Niv4, Niv5, Niv6, Niv7, Niv8 };
            var classes = new string?[] { Classe1, Classe2, Classe3, Classe4, Classe5, Classe6, Classe7, Classe8 };
            var maxLevel = levels.Max();
            var topClasse = classes[levels.ToList().IndexOf(maxLevel)];
            return topClasse is null ? "" : topClasse;
        }

        public bool ShouldBeNextLevel() => Xp >= NiveauSuivant;

        private short? _age;

        private string? _alignement;

        private string? _alphabet;

        private string? _archetype;

        private string? _artisanat1;

        private string? _artisanat2;

        private string? _artisanat3;

        private string? _backGround;

        private short? _beaute;

        private short? _charme;

        private string? _cheveux;

        private string? _classe1;

        private string? _classe2;

        private string? _classe3;

        private string? _classe4;

        private string? _classe5;

        private string? _classe6;

        private string? _classe7;

        private string? _classe8;

        private string? _dieu;

        private string? _domaine1;

        private string? _domaine2;

        private string? _domaine3;

        private string? _domaine4;

        private string? _dragonTotem;

        private string? _ecoleInterdite1;

        private string? _ecoleInterdite2;

        private string? _ecoleInterdite3;

        private string? _ecoleInterdite4;

        private string? _ecoleSpecialisation;

        private string? _elementShugenja;

        private string? _elementWujen;

        private short? _endurance = 0;

        private string? _energieElu1;

        private string? _energieElu2;

        private string? _energieElu3;

        private string? _energieSorcier1;

        private string? _energieSorcier2;

        private string? _ennemisJures;

        private short? _équilibre = 0;

        private short? _érudition = 0;

        private bool _exclu;

        private double _gainXp;

        private string? _histoire;

        private string? _image;

        private short? _intellect = 0;

        private short? _intuition = 0;

        private bool _isBelowNormalXpLevel;

        private string? _langueConnue;

        private short? _magnétisme = 0;

        private short? _maluxXp = 0;

        private short? _niv1 = 0;

        private short? _niv2 = 0;

        private short? _niv3 = 0;

        private short? _niv4 = 0;

        private short? _niv5 = 0;

        private short? _niv6 = 0;

        private short? _niv7 = 0;

        private short? _niv8 = 0;

        private int _niveau;

        private double _niveauGE;

        private double _nivSuivant;

        private string? _nom;

        private string? _ordreShugenja;

        private int _part = 1;

        private double _poids;

        private short? _précision = 0;

        private string? _profession1;

        private string? _profession2;

        private string? _profession3;

        private string? _profession4;

        private string? _profil;

        private short? _puissance;

        private string? _race;

        private short? _résistance;

        private string? _sexe;

        private short? _taille;

        private string? _titre;

        private double _totalPo;

        private long? _totalXp;

        private short? _vitalité;

        private short? _volonté;

        private double _xp;

        private string? _yeux;

        public Personnage()
        {
        }

        [Column]
        public short? Age { get => _age; set { Set(nameof(Age), ref _age, value); } }

        [Column]
        public string? Alignement { get => _alignement; set { Set(nameof(Alignement), ref _alignement, value); } }

        [Column]
        public string? Alphabet { get => _alphabet; set { Set(nameof(Alphabet), ref _alphabet, value); } }

        [Column]
        public string? Archetype { get => _archetype; set { Set(nameof(Archetype), ref _archetype, value); } }

        [Column("Artisanat_1")]
        public string? Artisanat1 { get => _artisanat1; set { Set(nameof(Artisanat1), ref _artisanat1, value); } }

        [Column("Artisanat_2")]
        public string? Artisanat2 { get => _artisanat2; set { Set(nameof(Artisanat2), ref _artisanat2, value); } }

        [Column("Artisanat_3")]
        public string? Artisanat3 { get => _artisanat3; set { Set(nameof(Artisanat3), ref _artisanat3, value); } }

        [Column]
        public string? BackGround { get => _backGround; set { Set(nameof(BackGround), ref _backGround, value); } }

        [Column]
        public short? Beaute { get => _beaute; set { Set(nameof(Beaute), ref _beaute, value); } }

        [Column]
        public short? Charme { get => _charme; set { Set(nameof(Charme), ref _charme, value); } }

        [Column]
        public string? Cheveux { get => _cheveux; set { Set(nameof(Cheveux), ref _cheveux, value); } }

        [Column("Classe_1")]
        public string? Classe1 { get => _classe1; set { Set(nameof(Classe1), ref _classe1, value); } }

        [Column("Classe_2")]
        public string? Classe2 { get => _classe2; set { Set(nameof(Classe2), ref _classe2, value); } }

        [Column("Classe_3")]
        public string? Classe3 { get => _classe3; set { Set(nameof(Classe3), ref _classe3, value); } }

        [Column("Classe_4")]
        public string? Classe4 { get => _classe4; set { Set(nameof(Classe4), ref _classe4, value); } }

        [Column("Classe_5")]
        public string? Classe5 { get => _classe5; set { Set(nameof(Classe5), ref _classe5, value); } }

        [Column("Classe_6")]
        public string? Classe6 { get => _classe6; set { Set(nameof(Classe6), ref _classe6, value); } }

        [Column("Classe_7")]
        public string? Classe7 { get => _classe7; set { Set(nameof(Classe7), ref _classe7, value); } }

        [Column("Classe_8")]
        public string? Classe8 { get => _classe8; set { Set(nameof(Classe8), ref _classe8, value); } }

        [Ignore]
        public ObservableCollection<DiceRoll> DiceRollEntries { get; private set; } = new ObservableCollection<DiceRoll>();

        [Column("dieu")]
        public string? Dieu { get => _dieu; set { Set(nameof(Dieu), ref _dieu, value); } }

        [Column("Domaine_1")]
        public string? Domaine1 { get => _domaine1; set { Set(nameof(Domaine1), ref _domaine1, value); } }

        [Column("Domaine_2")]
        public string? Domaine2 { get => _domaine2; set { Set(nameof(Domaine2), ref _domaine2, value); } }

        [Column("Domaine_3")]
        public string? Domaine3 { get => _domaine3; set { Set(nameof(Domaine3), ref _domaine3, value); } }

        [Column("Domaine_4")]
        public string? Domaine4 { get => _domaine4; set { Set(nameof(Domaine4), ref _domaine4, value); } }

        public string? DragonTotem { get => _dragonTotem; set { Set(nameof(DragonTotem), ref _dragonTotem, value); } }

        [Column("Ecole_interdite_1")]
        public string? EcoleInterdite1 { get => _ecoleInterdite1; set { Set(nameof(EcoleInterdite1), ref _ecoleInterdite1, value); } }

        [Column("Ecole_interdite_2")]
        public string? EcoleInterdite2 { get => _ecoleInterdite2; set { Set(nameof(EcoleInterdite2), ref _ecoleInterdite2, value); } }

        [Column("Ecole_interdite_3")]
        public string? EcoleInterdite3 { get => _ecoleInterdite3; set { Set(nameof(EcoleInterdite3), ref _ecoleInterdite3, value); } }

        [Column("Ecole_interdite_4")]
        public string? EcoleInterdite4 { get => _ecoleInterdite4; set { Set(nameof(EcoleInterdite4), ref _ecoleInterdite4, value); } }

        [Column]
        public string? EcoleSpecialisation { get => _ecoleSpecialisation; set { Set(nameof(EcoleSpecialisation), ref _ecoleSpecialisation, value); } }

        [Column]
        public string? ElementShugenja { get => _elementShugenja; set { Set(nameof(ElementShugenja), ref _elementShugenja, value); } }

        [Column]
        public string? ElementWujen { get => _elementWujen; set { Set(nameof(ElementWujen), ref _elementWujen, value); } }

        [Column]
        public short? Endurance { get => _endurance; set { Set(nameof(Endurance), ref _endurance, value); } }

        [Column("Energie_Elu_1")]
        public string? EnergieElu1 { get => _energieElu1; set { Set(nameof(EnergieElu1), ref _energieElu1, value); } }

        [Column("Energie_Elu_2")]
        public string? EnergieElu2 { get => _energieElu2; set { Set(nameof(EnergieElu2), ref _energieElu2, value); } }

        [Column("Energie_Elu_3")]
        public string? EnergieElu3 { get => _energieElu3; set { Set(nameof(EnergieElu3), ref _energieElu3, value); } }

        [Column("Energie_Sorcier_1")]
        public string? EnergieSorcier1 { get => _energieSorcier1; set { Set(nameof(EnergieSorcier1), ref _energieSorcier1, value); } }

        [Column("Energie_Sorcier_2")]
        public string? EnergieSorcier2 { get => _energieSorcier2; set { Set(nameof(EnergieSorcier2), ref _energieSorcier2, value); } }

        [Column]
        public string? EnnemisJures { get => _ennemisJures; set { Set(nameof(EnnemisJures), ref _ennemisJures, value); } }

        [Column]
        public short? Équilibre { get => _équilibre; set { Set(nameof(Équilibre), ref _équilibre, value); } }

        [Ignore]
        public ObservableCollection<Equipement> Equipement { get; private set; } = new ObservableCollection<Equipement>();

        [Column]
        public short? Érudition { get => _érudition; set { Set(nameof(Érudition), ref _érudition, value); } }

        [Column("exclu")]
        public bool Exclu { get => _exclu; set { Set(nameof(Exclu), ref _exclu, value); } }

        [Ignore]
        public double GainXp { get => _gainXp; set { Set(nameof(GainXp), ref _gainXp, value); } }

        [Column]
        public string? Histoire { get => _histoire; set { Set(nameof(Histoire), ref _histoire, value); } }

        [Column]
        public string? Image { get => _image; set { Set(nameof(Image), ref _image, value); } }

        [Column]
        public short? Intellect { get => _intellect; set { Set(nameof(Intellect), ref _intellect, value); } }

        [Column]
        public short? Intuition { get => _intuition; set { Set(nameof(Intuition), ref _intuition, value); } }

        [Column]
        public bool IsBelowNormalXpLevel { get => _isBelowNormalXpLevel; set { Set(nameof(IsBelowNormalXpLevel), ref _isBelowNormalXpLevel, value); } }

        [Column]
        public string? LangueConnue { get => _langueConnue; set { Set(nameof(LangueConnue), ref _langueConnue, value); } }

        [Column]
        public short? Magnétisme { get => _magnétisme; set { Set(nameof(Magnétisme), ref _magnétisme, value); } }

        [Column("MalusXP")]
        public short? MalusXp { get => _maluxXp; set { Set(nameof(MalusXp), ref _maluxXp, value); } }

        [Column("Niv_1")]
        public short? Niv1 { get => _niv1; set { Set(nameof(Niv1), ref _niv1, value); } }

        [Column("Niv_2")]
        public short? Niv2 { get => _niv2; set { Set(nameof(Niv2), ref _niv2, value); } }

        [Column("Niv_3")]
        public short? Niv3 { get => _niv3; set { Set(nameof(Niv3), ref _niv3, value); } }

        [Column("Niv_4")]
        public short? Niv4 { get => _niv4; set { Set(nameof(Niv4), ref _niv4, value); } }

        [Column("Niv_5")]
        public short? Niv5 { get => _niv5; set { Set(nameof(Niv5), ref _niv5, value); } }

        [Column("Niv_6")]
        public short? Niv6 { get => _niv6; set { Set(nameof(Niv6), ref _niv6, value); } }

        [Column("Niv_7")]
        public short? Niv7 { get => _niv7; set { Set(nameof(Niv7), ref _niv7, value); } }

        [Column("Niv_8")]
        public short? Niv8 { get => _niv8; set { Set(nameof(Niv8), ref _niv8, value); } }

        [Ignore]
        public int Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }

        [Ignore]
        public double NiveauGE { get => _niveauGE; set { Set(nameof(NiveauGE), ref _niveauGE, value); } }

        [Ignore]
        public double NiveauSuivant { get => _nivSuivant; set { Set(nameof(NiveauSuivant), ref _nivSuivant, value); } }

        [Column]
        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        [Column]
        public string? OrdreShugenja { get => _ordreShugenja; set { Set(nameof(OrdreShugenja), ref _ordreShugenja, value); } }

        [Ignore]
        public int Part { get => _part; set { Set(nameof(Part), ref _part, value); } }

        [Ignore]
        public ObservableCollection<PersonnageProgression> PersonnageProgression { get; private set; } = new ObservableCollection<PersonnageProgression>();

        [Column]
        public double Poids { get => _poids; set { Set(nameof(Poids), ref _poids, value); } }

        [Column]
        public short? Précision { get => _précision; set { Set(nameof(Précision), ref _précision, value); } }

        [Column("Profession_1")]
        public string? Profession1 { get => _profession1; set { Set(nameof(Profession1), ref _profession1, value); } }

        [Column("Profession_2")]
        public string? Profession2 { get => _profession2; set { Set(nameof(Profession2), ref _profession2, value); } }

        [Column("Profession_3")]
        public string? Profession3 { get => _profession3; set { Set(nameof(Profession3), ref _profession3, value); } }

        [Column("Profession_4")]
        public string? Profession4 { get => _profession4; set { Set(nameof(Profession4), ref _profession4, value); } }

        [Column("profil")]
        public string? Profil { get => _profil; set { Set(nameof(Profil), ref _profil, value); } }

        [Column]
        public short? Puissance { get => _puissance; set { Set(nameof(Puissance), ref _puissance, value); } }

        [Column]
        public string? Race { get => _race; set { Set(nameof(Race), ref _race, value); } }

        [Column]
        public short? Résistance { get => _résistance; set { Set(nameof(Résistance), ref _résistance, value); } }

        [Column]
        public string? Sexe { get => _sexe; set { Set(nameof(Sexe), ref _sexe, value); } }

        [Ignore]
        public ObservableCollection<SortPersonnage> SortPersonnage { get; private set; } = new ObservableCollection<SortPersonnage>();

        [Column]
        public short? Taille { get => _taille; set { Set(nameof(Taille), ref _taille, value); } }

        [Column("titre")]
        public string? Titre { get => _titre; set { Set(nameof(Titre), ref _titre, value); } }

        [Column]
        public double TotalPo { get => _totalPo; set { Set(nameof(TotalPo), ref _totalPo, value); } }

        [Column("TotalXP")]
        public long? TotalXp { get => _totalXp; set { Set(nameof(TotalXp), ref _totalXp, value); } }

        [Column]
        public short? Vitalité { get => _vitalité; set { Set(nameof(Vitalité), ref _vitalité, value); } }

        [Column]
        public short? Volonté { get => _volonté; set { Set(nameof(Volonté), ref _volonté, value); } }

        [Ignore]
        public double Xp { get => _xp; set { Set(nameof(Xp), ref _xp, value); } }

        [Column]
        public string? Yeux { get => _yeux; set { Set(nameof(Yeux), ref _yeux, value); } }
    }
}