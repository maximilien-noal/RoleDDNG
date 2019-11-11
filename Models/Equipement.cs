namespace RoleDDNG.Models
{
    public class Equipement
    {
        public int Id { get; set; }
        public string NomObjet { get; set; }
        public string Personnage { get; set; }
        public short? Place { get; set; }
        public short? Pos { get; set; }
        public string Type { get; set; }
        public short? Charge { get; set; }
        public bool SurPersonnage { get; set; }
        public short? Multiple { get; set; }

        public virtual Personnage PersonnageNavigation { get; set; }
    }
}