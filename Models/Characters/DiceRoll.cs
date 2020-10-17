using GalaSoft.MvvmLight;

using PetaPoco;

using System;

namespace RoleDDNG.Models.Characters
{
    [TableName(nameof(DiceRoll))]
    [PrimaryKey(nameof(Id))]
    public class DiceRoll : ObservableObject
    {
        private int _id;

        private Personnage? _character;

        private DateTime _dateTime = DateTime.Now;

        private int _dices = 1;

        private int _sides = 6;

        private int _sum;

        private int _result;

        public DiceRoll()
        {
        }

        public int Id { get => _id; set { Set(nameof(Id), ref _id, value); } }

        public Personnage? Character { get => _character; set { Set(nameof(Character), ref _character, value); } }

        public int Result { get => _result; set { Set(nameof(Result), ref _result, value); } }

        [Column("DDateTime")]
        public DateTime DateTime { get => _dateTime; set { Set(nameof(DateTime), ref _dateTime, value); } }

        public int Dices { get => _dices; set { Set(nameof(Dices), ref _dices, value); } }

        public int Sides { get => _sides; set { Set(nameof(Sides), ref _sides, value); } }

        [Column("DSum")]
        public int Sum { get => _sum; set { Set(nameof(Sum), ref _sum, value); } }
    }
}