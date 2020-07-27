using GalaSoft.MvvmLight;
using System;

namespace RoleDDNG.Models.Characters
{
    public class DiceRollEntry : ObservableObject
    {
        private DateTime _dateTime = DateTime.Now;

        private int _dices = 1;

        private int _sides = 6;

        private int _sum = 0;

        public DiceRollEntry()
        {
        }

        public DateTime DateTime { get => _dateTime; set { Set(nameof(DateTime), ref _dateTime, value); } }

        public int Dices { get => _dices; set { Set(nameof(Dices), ref _dices, value); } }

        public int Sides { get => _sides; set { Set(nameof(Sides), ref _sides, value); } }

        public int Sum { get => _sum; set { Set(nameof(Sum), ref _sum, value); } }
    }
}