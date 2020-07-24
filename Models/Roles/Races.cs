using GalaSoft.MvvmLight;
using System.ComponentModel.DataAnnotations.Schema;

namespace RoleDDNG.Models.Roles
{
    [Table("Race")]
    public class Races : ObservableObject
    {
        private short? _adjNiv = 0;

        private string _race = "";

        private string _source = "";

        public short? AdjNiv { get => _adjNiv; set { Set(nameof(AdjNiv), ref _adjNiv, value); } }

        /// <summary> Key </summary>
        public string Race { get => _race; set { Set(nameof(Race), ref _race, value); } }

        public string Source { get => _source; set { Set(nameof(Source), ref _source, value); } }
    }
}