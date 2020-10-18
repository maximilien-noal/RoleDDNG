using PetaPoco;

using System;

namespace RoleDDNG.Models.Characters
{
    [TableName(nameof(DiceRoll))]
    [PrimaryKey(nameof(Id), AutoIncrement = true)]
    public class DiceRoll : ObservableObject
    {
        private int _id;

        private string _character = "";

        private DateTime _dateTime = DateTime.Now.WithoutMs();

        private int _dices = 1;

        private int _sides = 6;

        private int _sum;

        private string _results = "";

        public DiceRoll()
        {
        }

        [Column]
        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }

        [Column("Personnage")]
        public string Character { get => _character; set { Set(nameof(Character), ref _character, value); } }

        [Column]
        public string Results { get => _results; set { Set(nameof(Results), ref _results, value); } }

        [Column("DDateTime")]
        public DateTime DateTime { get => _dateTime; set { Set(nameof(DateTime), ref _dateTime, value); } }

        [Column]
        public int Dices { get => _dices; set { Set(nameof(Dices), ref _dices, value); } }

        [Column]
        public int Sides { get => _sides; set { Set(nameof(Sides), ref _sides, value); } }

        [Column("DSum")]
        public int Sum { get => _sum; set { Set(nameof(Sum), ref _sum, value); } }
    }
}