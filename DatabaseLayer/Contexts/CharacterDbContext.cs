using System;

using EntityFrameworkCore.Jet;

using Microsoft.EntityFrameworkCore;

using RoleDDNG.Models.Characters;

namespace RoleDDNG.DatabaseLayer.Contexts
{
    public class CharacterDbContext : DbContext
    {
        private string _dbFilePath = "";

        public CharacterDbContext(string dbFilePath)
        {
            _dbFilePath = dbFilePath;
        }

        public CharacterDbContext(DbContextOptions<CharacterDbContext> options)
            : base(options)
        {
        }

        public virtual DbSet<Objets> Objets { get; set; }

        public virtual DbSet<Personnage> Personnage { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (optionsBuilder is null)
            {
                throw new ArgumentNullException(nameof(optionsBuilder));
            }

            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseJet($"Provider=Microsoft.ACE.OLEDB.12.0;{_dbFilePath}");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            if (modelBuilder is null)
            {
                throw new ArgumentNullException(nameof(modelBuilder));
            }

            modelBuilder.HasAnnotation("ProductVersion", "2.2.6-servicing-10079");

            modelBuilder.Entity<Objets>(entity =>
            {
                entity.HasKey(e => e.NomObjet)
                    .HasName("PrimaryKey");

                entity.Property(e => e.NomObjet)
                    .HasColumnName("nomObjet")
                    .HasMaxLength(50)
                    .ValueGeneratedNever();

                entity.Property(e => e.DegatDes).HasMaxLength(10);

                entity.Property(e => e.GroupeObjet).HasMaxLength(50);

                entity.Property(e => e.PoidsBase).HasDefaultValueSql("0");

                entity.Property(e => e.Taille).HasMaxLength(50);

                entity.Property(e => e.Type).HasMaxLength(50);

                entity.Property(e => e.TypeCa).HasMaxLength(15);
            });

            modelBuilder.Entity<Personnage>(entity =>
            {
                entity.HasKey(e => e.Nom)
                    .HasName("PrimaryKey");

                entity.Property(e => e.Nom)
                    .HasMaxLength(50)
                    .ValueGeneratedNever();

                entity.Property(e => e.Age).HasDefaultValueSql("0");

                entity.Property(e => e.Alignement).HasMaxLength(30);

                entity.Property(e => e.Alphabet).HasMaxLength(255);

                entity.Property(e => e.Archetype).HasMaxLength(50);

                entity.Property(e => e.Artisanat1)
                    .HasColumnName("Artisanat_1")
                    .HasMaxLength(50);

                entity.Property(e => e.Artisanat2)
                    .HasColumnName("Artisanat_2")
                    .HasMaxLength(50);

                entity.Property(e => e.Artisanat3)
                    .HasColumnName("Artisanat_3")
                    .HasMaxLength(50);

                entity.Property(e => e.Beaute).HasDefaultValueSql("0");

                entity.Property(e => e.Charme).HasDefaultValueSql("0");

                entity.Property(e => e.Cheveux).HasMaxLength(15);

                entity.Property(e => e.Classe1)
                    .HasColumnName("Classe_1")
                    .HasMaxLength(30);

                entity.Property(e => e.Classe2)
                    .HasColumnName("Classe_2")
                    .HasMaxLength(30);

                entity.Property(e => e.Classe3)
                    .HasColumnName("Classe_3")
                    .HasMaxLength(30);

                entity.Property(e => e.Classe4)
                    .HasColumnName("Classe_4")
                    .HasMaxLength(50);

                entity.Property(e => e.Classe5)
                    .HasColumnName("Classe_5")
                    .HasMaxLength(50);

                entity.Property(e => e.Classe6)
                    .HasColumnName("Classe_6")
                    .HasMaxLength(50);

                entity.Property(e => e.Classe7)
                    .HasColumnName("Classe_7")
                    .HasMaxLength(50);

                entity.Property(e => e.Classe8)
                    .HasColumnName("Classe_8")
                    .HasMaxLength(50);

                entity.Property(e => e.Date).HasMaxLength(50);

                entity.Property(e => e.Dieu)
                    .HasColumnName("dieu")
                    .HasMaxLength(30);

                entity.Property(e => e.Domaine1)
                    .HasColumnName("Domaine_1")
                    .HasMaxLength(50);

                entity.Property(e => e.Domaine2)
                    .HasColumnName("Domaine_2")
                    .HasMaxLength(50);

                entity.Property(e => e.Domaine3)
                    .HasColumnName("Domaine_3")
                    .HasMaxLength(50);

                entity.Property(e => e.Domaine4)
                    .HasColumnName("Domaine_4")
                    .HasMaxLength(50);

                entity.Property(e => e.DragonTotem).HasMaxLength(50);

                entity.Property(e => e.EcoleInterdite1).HasColumnName("Ecole_interdite_1");

                entity.Property(e => e.EcoleInterdite2).HasColumnName("Ecole_interdite_2");

                entity.Property(e => e.EcoleInterdite3).HasColumnName("Ecole_interdite_3");

                entity.Property(e => e.EcoleInterdite4).HasColumnName("Ecole_interdite_4");

                entity.Property(e => e.EcoleSpecialisation).HasMaxLength(50);

                entity.Property(e => e.ElementShugenja).HasMaxLength(50);

                entity.Property(e => e.ElementWujen).HasMaxLength(50);

                entity.Property(e => e.Endurance).HasDefaultValueSql("0");

                entity.Property(e => e.EnergieElu1)
                    .HasColumnName("EnergieElu_1")
                    .HasMaxLength(50);

                entity.Property(e => e.EnergieElu2)
                    .HasColumnName("EnergieElu_2")
                    .HasMaxLength(50);

                entity.Property(e => e.EnergieElu3)
                    .HasColumnName("EnergieElu_3")
                    .HasMaxLength(50);

                entity.Property(e => e.EnergieSorcier1)
                    .HasColumnName("EnergieSorcier_1")
                    .HasMaxLength(50);

                entity.Property(e => e.EnergieSorcier2)
                    .HasColumnName("EnergieSorcier_2")
                    .HasMaxLength(50);

                entity.Property(e => e.EnnemisJures).HasMaxLength(255);

                entity.Property(e => e.Exclu)
                    .HasColumnName("exclu")
                    .HasColumnType("bit");

                entity.Property(e => e.Image).HasMaxLength(255);

                entity.Property(e => e.Intellect).HasDefaultValueSql("0");

                entity.Property(e => e.Intuition).HasDefaultValueSql("0");

                entity.Property(e => e.LangueConnue).HasMaxLength(255);

                entity.Property(e => e.Magnétisme).HasDefaultValueSql("0");

                entity.Property(e => e.MalusXp)
                    .HasColumnName("MalusXP")
                    .HasDefaultValueSql("0");

                entity.Property(e => e.Niv1)
                    .HasColumnName("Niv_1")
                    .HasDefaultValueSql("0");

                entity.Property(e => e.Niv2)
                    .HasColumnName("Niv_2")
                    .HasDefaultValueSql("0");

                entity.Property(e => e.Niv3)
                    .HasColumnName("Niv_3")
                    .HasDefaultValueSql("0");

                entity.Property(e => e.Niv4).HasColumnName("Niv_4");

                entity.Property(e => e.Niv5).HasColumnName("Niv_5");

                entity.Property(e => e.Niv6).HasColumnName("Niv_6");

                entity.Property(e => e.Niv7).HasColumnName("Niv_7");

                entity.Property(e => e.Niv8).HasColumnName("Niv_8");

                entity.Property(e => e.OrdreShugenja).HasMaxLength(50);

                entity.Property(e => e.Profession1)
                    .HasColumnName("Profession_1")
                    .HasMaxLength(50);

                entity.Property(e => e.Profession2)
                    .HasColumnName("Profession_2")
                    .HasMaxLength(50);

                entity.Property(e => e.Profession3)
                    .HasColumnName("Profession_3")
                    .HasMaxLength(50);

                entity.Property(e => e.Profession4)
                    .HasColumnName("Profession_4")
                    .HasMaxLength(50);

                entity.Property(e => e.Profil)
                    .HasColumnName("profil")
                    .HasMaxLength(30);

                entity.Property(e => e.Précision).HasDefaultValueSql("0");

                entity.Property(e => e.Puissance).HasDefaultValueSql("0");

                entity.Property(e => e.Race).HasMaxLength(50);

                entity.Property(e => e.Résistance).HasDefaultValueSql("0");

                entity.Property(e => e.Sexe).HasMaxLength(10);

                entity.Property(e => e.Taille).HasDefaultValueSql("0");

                entity.Property(e => e.Titre)
                    .HasColumnName("titre")
                    .HasMaxLength(30);

                entity.Property(e => e.TotalXp).HasColumnName("TotalXP");

                entity.Property(e => e.Vitalité).HasDefaultValueSql("0");

                entity.Property(e => e.Volonté).HasDefaultValueSql("0");

                entity.Property(e => e.Yeux).HasMaxLength(15);

                entity.Property(e => e.Équilibre).HasDefaultValueSql("0");

                entity.Property(e => e.Érudition).HasDefaultValueSql("0");
            });
        }
    }
}