using System;
using System.Collections.Generic;

namespace MdbCli
{
    public class SortPersonnage
    {
        public int Id { get; set; }
        public string NomPerso { get; set; }
        public string ClassePerso { get; set; }
        public string NomSort { get; set; }
        public string Niveau { get; set; }

        public virtual Personnage NomPersoNavigation { get; set; }
    }
}
