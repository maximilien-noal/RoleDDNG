using GalaSoft.MvvmLight;
using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [TableName(nameof(Race))]
    [PrimaryKey("race")]
    public class RacePersonnage : ObservableObject
    {
        private int? _adjNiv = 0;

        private string? _description;

        private string? _race;

        public RacePersonnage()
        {
        }

        [Column(nameof(AdjNiv))]
        public int? AdjNiv { get => _adjNiv; set { Set(nameof(AdjNiv), ref _adjNiv, value); } }

        [Column]
        public string? Description { get => _description; set { Set(nameof(Description), ref _description, value); } }

        [Column("race")]
        public string? Race { get => _race; set { Set(nameof(Race), ref _race, value); } }
    }
}