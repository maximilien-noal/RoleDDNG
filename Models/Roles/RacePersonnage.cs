using GalaSoft.MvvmLight;
using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    [TableName(nameof(Race))]
    [PrimaryKey("race")]
    public class RacePersonnage : ObservableObject
    {
        private int? _adjNiv = 0;

        private string _race = "";

        private string _source = "";

        public RacePersonnage()
        {
        }

        [Column(nameof(AdjNiv))]
        public int? AdjNiv { get => _adjNiv; set { Set(nameof(AdjNiv), ref _adjNiv, value); } }

        [Column("race")]
        public string Race { get => _race; set { Set(nameof(Race), ref _race, value); } }

        [Column]
        public string Source { get => _source; set { Set(nameof(Source), ref _source, value); } }
    }
}