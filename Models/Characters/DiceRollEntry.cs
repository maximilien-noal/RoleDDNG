using System;
using System.Collections.Generic;

namespace RoleDDNG.Models.Characters
{
    public class DiceRollEntry
    {
        public DateTime DateAndTime { get; set; } = DateTime.MinValue;

        public int NumberOfDices { get; set; } = 1;

        public int NumberofSides { get; set; } = 6;

        public List<int> Results { get; private set; } = new List<int>();
    }
}