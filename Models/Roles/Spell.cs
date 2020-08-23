using GalaSoft.MvvmLight;
using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [TableName("Sort")]
    [PrimaryKey(nameof(Nom))]
    public class Spell : ObservableObject
    {
        private string? _cible;

        private string? _classeNiveau;

        private string? _composantes;

        private short? _ddDeDefficultee = 0;

        private string? _duree;

        private string? _ecole;

        private string? _effet;

        private int? _epique = 0;

        private string? _explication;

        private int? _id = 0;

        private string? _jetDeSauvegarde;

        private string? _nom;

        private string? _portee;

        private string? _pourDevelopper;

        private string? _resistanceALaMagie;

        private string? _source;

        private string? _tempsIncantation;

        private string? _version;

        private string? _zoneDeffet;

        public Spell()
        {
        }

        public string? Cible { get => _cible; set { Set(nameof(_cible), ref _cible, value); } }

        [Column("Classe/Niveau")]
        public string? ClasseNiveau { get => _classeNiveau; set { Set(nameof(ClasseNiveau), ref _classeNiveau, value); } }

        public string? Composantes { get => _composantes; set { Set(nameof(Composantes), ref _composantes, value); } }

        public short? DDdeDifficultee { get => _ddDeDefficultee; set { Set(nameof(DDdeDifficultee), ref _ddDeDefficultee, value); } }

        [Column("Durée")]
        public string? Duree { get => _duree; set { Set(nameof(Duree), ref _duree, value); } }

        public string? Ecole { get => _ecole; set { Set(nameof(Ecole), ref _ecole, value); } }

        public string? Effet { get => _effet; set { Set(nameof(Effet), ref _effet, value); } }

        public int? Epique { get => _epique; set { Set(nameof(Epique), ref _epique, value); } }

        public string? Explication { get => _explication; set { Set(nameof(Explication), ref _explication, value); } }

        public int? Id { get => _id; set { Set(nameof(Id), ref _id, value); } }

        [Column("Jet de sauvegarde")]
        public string? JetDeSauvegarde { get => _jetDeSauvegarde; set { Set(nameof(JetDeSauvegarde), ref _jetDeSauvegarde, value); } }

        public string? Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        [Column("Portée")]
        public string? Portee { get => _portee; set { Set(nameof(Portee), ref _portee, value); } }

        [Column("PourDevellopper")]
        public string? PourDevelopper { get => _pourDevelopper; set { Set(nameof(PourDevelopper), ref _pourDevelopper, value); } }

        [Column("Résistance à la magie")]
        public string? ResistanceALaMagie { get => _resistanceALaMagie; set { Set(nameof(ResistanceALaMagie), ref _resistanceALaMagie, value); } }

        [Column("source")]
        public string? Source { get => _source; set { Set(nameof(Source), ref _source, value); } }

        [Column("Temps d’incantation")]
        public string? TempsIncantation { get => _tempsIncantation; set { Set(nameof(TempsIncantation), ref _tempsIncantation, value); } }

        [Column("version")]
        public string? Version { get => _version; set { Set(nameof(Version), ref _version, value); } }

        [Column("Zone d'effet")]
        public string? ZoneDeffet { get => _zoneDeffet; set { Set(nameof(_zoneDeffet), ref _zoneDeffet, value); } }
    }
}