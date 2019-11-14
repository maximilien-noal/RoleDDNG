using System;
using System.Collections.Generic;
using System.Text;

namespace RoleDDNG.Models
{
    public class DiceRollEntry
    {
        public DateTime DateAndTime { get; set; } = DateTime.MinValue;

        public int NumberofSides { get; set; } = 6;

        public int NumberOfDices { get; set; } = 1;

        public List<int> Results { get; private set; } = new List<int>();
    }
}