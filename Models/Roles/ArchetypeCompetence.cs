using GalaSoft.MvvmLight;

using System.ComponentModel.DataAnnotations.Schema;

namespace RoleDDNG.Models.Roles
{
    public class ArchetypeCompetence : ObservableObject
    {
        private string _archetype = "";

        public string Archetype { get => _archetype; set { Set(nameof(Archetype), ref _archetype, value); } }

        private string _competence = "";

        [Column("competence")]
        public string Competence { get => _competence; set { Set(nameof(Competence), ref _competence, value); } }

        private int _modif = 0;

        [Column("modif")]
        public int Modif { get => _modif; set { Set(nameof(Modif), ref _modif, value); } }

        private string _texte = "";

        public string Texte { get => _texte; set { Set(nameof(Texte), ref _texte, value); } }
    }
}