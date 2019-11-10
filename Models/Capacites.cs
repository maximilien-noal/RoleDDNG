using System;
using System.Collections.Generic;

namespace RoleDDNG.Models
{
    public class Capacites
    {
        public string Nom { get; set; }
        public string Personnage { get; set; }
        public string Type { get; set; }
        public short? Modificateur1 { get; set; }
        public short? Modificateur2 { get; set; }
        public string Texte { get; set; }
    }
}
