using PetaPoco;

namespace RoleDDNG.Models.Characters
{
    public class ObjetsPropriete : ObservableObject
    {
        private string? _nomObjet;

        private string? _propriete1;

        private string? _propriete2;

        private string? _propriete3;

        private string? _valeur;

        public ObjetsPropriete()
        {
        }

        [Column("nomObjet")]
        public string? NomObjet { get => _nomObjet; set { Set(nameof(NomObjet), ref _nomObjet, value); } }

        [Column("Propriete_1")]
        public string? Propriete1 { get => _propriete1; set { Set(nameof(Propriete1), ref _propriete1, value); } }

        [Column("Propriete_2")]
        public string? Propriete2 { get => _propriete2; set { Set(nameof(Propriete2), ref _propriete2, value); } }

        [Column("Propriete_3")]
        public string? Propriete3 { get => _propriete3; set { Set(nameof(Propriete3), ref _propriete3, value); } }

        [Column("valeur")]
        public string? Valeur { get => _valeur; set { Set(nameof(Valeur), ref _valeur, value); } }
    }
}