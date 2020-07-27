using PetaPoco;

using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Roles
{
    public class ArchetypeCompetence : ObservableObject
    {
        private string _archetype = "";

        private string _competence = "";

        private int _modif = 0;

        private string _texte = "";

        public ArchetypeCompetence()
        {
        }

        public string Archetype { get => _archetype; set { Set(nameof(Archetype), ref _archetype, value); } }

        [Column("competence")]
        public string Competence { get => _competence; set { Set(nameof(Competence), ref _competence, value); } }

        [Column("modif")]
        public int Modif { get => _modif; set { Set(nameof(Modif), ref _modif, value); } }

        public string Texte { get => _texte; set { Set(nameof(Texte), ref _texte, value); } }
    }
}