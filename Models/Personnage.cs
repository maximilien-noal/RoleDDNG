using System;
using System.Collections.Generic;

namespace MdbCli
{
    public class Personnage
    {
        public Personnage()
        {
            Equipement = new HashSet<Equipement>();
            PersonnageProgression = new HashSet<PersonnageProgression>();
            SortPersonnage = new HashSet<SortPersonnage>();
        }

        public string Nom { get; set; }
        public string Race { get; set; }
        public string Classe1 { get; set; }
        public short? Niv1 { get; set; }
        public string Classe2 { get; set; }
        public short? Niv2 { get; set; }
        public string Classe3 { get; set; }
        public short? Niv3 { get; set; }
        public short? Endurance { get; set; }
        public short? Puissance { get; set; }
        public short? Précision { get; set; }
        public short? Équilibre { get; set; }
        public short? Résistance { get; set; }
        public short? Vitalité { get; set; }
        public short? Intellect { get; set; }
        public short? Érudition { get; set; }
        public short? Intuition { get; set; }
        public short? Volonté { get; set; }
        public short? Magnétisme { get; set; }
        public short? Charme { get; set; }
        public string Profil { get; set; }
        public string Titre { get; set; }
        public string Dieu { get; set; }
        public string Alignement { get; set; }
        public short? Beaute { get; set; }
        public string Cheveux { get; set; }
        public string Yeux { get; set; }
        public string Sexe { get; set; }
        public double? Poids { get; set; }
        public short? Taille { get; set; }
        public string Date { get; set; }
        public short? Age { get; set; }
        public short? MalusXp { get; set; }
        public string EcoleSpecialisation { get; set; }
        public string LangueConnue { get; set; }
        public string Alphabet { get; set; }
        public bool Exclu { get; set; }
        public string Domaine1 { get; set; }
        public string Domaine2 { get; set; }
        public string Domaine3 { get; set; }
        public string BackGround { get; set; }
        public string Histoire { get; set; }
        public string EcoleInterdite1 { get; set; }
        public string EcoleInterdite2 { get; set; }
        public string Classe4 { get; set; }
        public string Classe5 { get; set; }
        public string Classe6 { get; set; }
        public short? Niv4 { get; set; }
        public short? Niv5 { get; set; }
        public short? Niv6 { get; set; }
        public string Archetype { get; set; }
        public double? TotalPo { get; set; }
        public int? TotalXp { get; set; }
        public string Profession1 { get; set; }
        public string Profession2 { get; set; }
        public string Profession3 { get; set; }
        public string Profession4 { get; set; }
        public string Artisanat1 { get; set; }
        public string Artisanat2 { get; set; }
        public string Artisanat3 { get; set; }
        public string EnergieElu1 { get; set; }
        public string EnergieElu2 { get; set; }
        public string EnergieElu3 { get; set; }
        public string EnergieSorcier1 { get; set; }
        public string EnergieSorcier2 { get; set; }
        public string ElementWujen { get; set; }
        public string EnnemisJures { get; set; }
        public string Image { get; set; }
        public string Domaine4 { get; set; }
        public string EcoleInterdite3 { get; set; }
        public string EcoleInterdite4 { get; set; }
        public string Classe7 { get; set; }
        public string Classe8 { get; set; }
        public short? Niv7 { get; set; }
        public short? Niv8 { get; set; }
        public string DragonTotem { get; set; }
        public string ElementShugenja { get; set; }
        public string OrdreShugenja { get; set; }

        public virtual ICollection<Equipement> Equipement { get; set; }
        public virtual ICollection<PersonnageProgression> PersonnageProgression { get; set; }
        public virtual ICollection<SortPersonnage> SortPersonnage { get; set; }
    }
}
