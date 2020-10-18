using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [TableName(nameof(Classe))]
    [PrimaryKey(nameof(Classe))]
    public class ClassePersonnage : ObservableObject
    {
        private bool _affichage;

        private string? _alignements;

        private string? _anciens;

        private string? _armesArmures;

        private string? _caracSortBonus;

        private string? _caracSortDD;

        private string? _caracSortNiv;

        private string? _classe;

        private int _competenceNiveauSupp;

        private string? _conditions;

        private string? _description;

        private bool _estLanceurSort;

        private bool _estLimite;

        private bool _estTout;

        private string? _listSort;

        private int _nls;

        private int _number;

        private int _pointDeVie;

        private int _progression;

        private string? _sorts;

        private string? _type;

        public ClassePersonnage()
        {
        }

        [Column]
        public bool Affichage { get => _affichage; set { Set(nameof(Affichage), ref _affichage, value); } }

        [Column]
        public string? Alignements { get => _alignements; set { Set(nameof(Alignements), ref _alignements, value); } }

        [Column]
        public string? Anciens { get => _anciens; set { Set(nameof(Anciens), ref _anciens, value); } }

        [Column("Armes_Armures")]
        public string? ArmesArmures { get => _armesArmures; set { Set(nameof(ArmesArmures), ref _armesArmures, value); } }

        public string? CaracSortBonus { get => _caracSortBonus; set { Set(nameof(CaracSortBonus), ref _caracSortBonus, value); } }

        [Column("CaracSort_DD")]
        public string? CaracSortDD { get => _caracSortDD; set { Set(nameof(CaracSortDD), ref _caracSortDD, value); } }

        [Column("CaracSort_Niv")]
        public string? CaracSortNiv { get => _caracSortNiv; set { Set(nameof(CaracSortNiv), ref _caracSortNiv, value); } }

        [Column]
        public string? Classe { get => _classe; set { Set(nameof(Classe), ref _classe, value); } }

        [Column]
        public int CompetenceNiveauSupp { get => _competenceNiveauSupp; set { Set(nameof(CompetenceNiveauSupp), ref _competenceNiveauSupp, value); } }

        [Column]
        public string? Conditions { get => _conditions; set { Set(nameof(Conditions), ref _conditions, value); } }

        [Column]
        public string? Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        [Column]
        public bool EstLanceurSort { get => _estLanceurSort; set { Set(nameof(EstLanceurSort), ref _estLanceurSort, value); } }

        [Column]
        public bool EstLimite { get => _estLimite; set { Set(nameof(EstLimite), ref _estLimite, value); } }

        [Column]
        public bool EstTout { get => _estTout; set { Set(nameof(EstTout), ref _estTout, value); } }

        [Column]
        public string? ListeSort { get => _listSort; set { Set(nameof(ListeSort), ref _listSort, value); } }

        [Column]
        public int NLS { get => _nls; set { Set(nameof(NLS), ref _nls, value); } }

        [Column("n°")]
        public int Number { get => _number; set { Set(nameof(Number), ref _number, value); } }

        [Column]
        public int PointDeVie { get => _pointDeVie; set { Set(nameof(PointDeVie), ref _pointDeVie, value); } }

        [Column]
        public int Progression { get => _progression; set { Set(nameof(Progression), ref _progression, value); } }

        [Column]
        public string? Sorts { get => _sorts; set { Set(nameof(Sorts), ref _sorts, value); } }

        [Column]
        public string? Type { get => _type; set { Set(nameof(Type), ref _type, value); } }
    }
}