using GalaSoft.MvvmLight;

using System.ComponentModel.DataAnnotations.Schema;

namespace RoleDDNG.Models.Roles
{
    [Table(nameof(Classe))]
    public class ClassePersonnage : ObservableObject
    {
        private string _classe = "";

        /// <summary>
        /// Key
        /// </summary>
        public string Classe { get => _classe; set { Set(nameof(Classe), ref _classe, value); } }

        private int _pointDeVie = 0;

        public int PointDeVie { get => _pointDeVie; set { Set(nameof(PointDeVie), ref _pointDeVie, value); } }

        private int _competenceNiveauSupp = 0;

        public int CompetenceNiveauSupp { get => _competenceNiveauSupp; set { Set(nameof(CompetenceNiveauSupp), ref _competenceNiveauSupp, value); } }

        private string _type = "";

        public string Type { get => _type; set { Set(nameof(Type), ref _type, value); } }

        private bool _estLanceurSort = false;

        public bool EstLanceurSort { get => _estLanceurSort; set { Set(nameof(EstLanceurSort), ref _estLanceurSort, value); } }

        private bool _estLimite = false;

        public bool EstLimite { get => _estLimite; set { Set(nameof(EstLimite), ref _estLimite, value); } }

        private bool _estTout = false;

        public bool EstTout { get => _estTout; set { Set(nameof(EstTout), ref _estTout, value); } }

        private int _nls = 0;

        public int NLS { get => _nls; set { Set(nameof(NLS), ref _nls, value); } }

        private int _progression = 0;

        public int Progression { get => _progression; set { Set(nameof(Progression), ref _progression, value); } }

        private string _caracSortBonus = "";

        public string CaracSortBonus { get => _caracSortBonus; set { Set(nameof(CaracSortBonus), ref _caracSortBonus, value); } }

        private string _caracSortDD = "";

        [Column("CaracSort_DD")]
        public string CaracSortDD { get => _caracSortDD; set { Set(nameof(CaracSortDD), ref _caracSortDD, value); } }

        private string _caracSortNiv = "";

        [Column("CaracSort_Niv")]
        public string CaracSortNiv { get => _caracSortNiv; set { Set(nameof(CaracSortNiv), ref _caracSortNiv, value); } }

        private string _listSort = "";

        public string ListeSort { get => _listSort; set { Set(nameof(ListeSort), ref _listSort, value); } }

        private string _alignements = "";

        public string Alignements { get => _alignements; set { Set(nameof(Alignements), ref _alignements, value); } }

        private string _conditions = "";

        public string Conditions { get => _conditions; set { Set(nameof(Conditions), ref _conditions, value); } }

        private string _armesArmures = "";

        [Column("Armes_Armures")]
        public string ArmesArmures { get => _armesArmures; set { Set(nameof(ArmesArmures), ref _armesArmures, value); } }

        private string _sorts = "";

        public string Sorts { get => _sorts; set { Set(nameof(Sorts), ref _sorts, value); } }

        private string _anciens = "";

        public string Anciens { get => _anciens; set { Set(nameof(Anciens), ref _anciens, value); } }

        private string _description = "";

        public string Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        private bool _affichage = false;

        public bool Affichage { get => _affichage; set { Set(nameof(Affichage), ref _affichage, value); } }

        private int _number = 0;

        [Column("n°")]
        public int Number { get => _number; set { Set(nameof(Number), ref _number, value); } }
    }
}