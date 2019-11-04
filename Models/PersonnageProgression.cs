using System;
using System.Collections.Generic;

namespace MdbCli
{
    public class PersonnageProgression
    {
        public int Id { get; set; }
        public string Nom { get; set; }
        public string Classe { get; set; }
        public short? Niveau { get; set; }
        public short? Pv { get; set; }
        public short? For { get; set; }
        public short? Dex { get; set; }
        public short? Con { get; set; }
        public short? Int { get; set; }
        public short? Sag { get; set; }
        public short? Cha { get; set; }
        public short? Niv { get; set; }
        public string Competence { get; set; }
        public short? Modifint { get; set; }

        public virtual Personnage NomNavigation { get; set; }
    }
}
