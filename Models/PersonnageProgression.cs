namespace RoleDDNG.Models
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
#pragma warning disable CA1720 // Identifier contains type name (Justification : Legacy column name)
        public short? Int { get; set; }
#pragma warning restore CA1720 // Identifier contains type name (Justification : Legacy column name)
        public short? Sag { get; set; }
        public short? Cha { get; set; }
        public short? Niv { get; set; }
        public string Competence { get; set; }
        public short? Modifint { get; set; }

        public virtual Personnage NomNavigation { get; set; }
    }
}