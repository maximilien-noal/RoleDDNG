using System;
using EntityFrameworkCore.Jet;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;
using RoleDDNG.Models;
using Version = RoleDDNG.Models.Version;

namespace RoleDDNG.MdbAccess
{
    internal class ModelContext : DbContext
    {
        public ModelContext()
        {
        }

        public ModelContext(DbContextOptions<ModelContext> options)
            : base(options)
        {
        }

        public virtual DbSet<Capacites> Capacites { get; set; }
        public virtual DbSet<Donsperso> Donsperso { get; set; }
        public virtual DbSet<Equipement> Equipement { get; set; }
        public virtual DbSet<Objets> Objets { get; set; }
        public virtual DbSet<ObjetsPropriete> ObjetsPropriete { get; set; }
        public virtual DbSet<Personnage> Personnage { get; set; }
        public virtual DbSet<PersonnageDons> PersonnageDons { get; set; }
        public virtual DbSet<PersonnageProgression> PersonnageProgression { get; set; }
        public virtual DbSet<SortListe> SortListe { get; set; }
        public virtual DbSet<SortPersonnage> SortPersonnage { get; set; }
        public virtual DbSet<Models.Version> Version { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if(optionsBuilder == null)
            {
                throw new ArgumentNullException(nameof(optionsBuilder));
            }
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseJet("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\max\\Desktop\\db1.accdb;");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            if(modelBuilder == null)
            {
                throw new ArgumentNullException(nameof(modelBuilder));
            }
            modelBuilder.HasAnnotation("ProductVersion", "2.2.6-servicing-10079");

            modelBuilder.Entity<Capacites>(entity =>
            {
                entity.HasKey(e => e.Nom)
                    .HasName("PrimaryKey");

                entity.Property(e => e.Nom)
                    .HasMaxLength(50)
                    .ValueGeneratedNever();

                entity.Property(e => e.Modificateur1).HasColumnName("Modificateur_1");

                entity.Property(e => e.Modificateur2).HasColumnName("Modificateur_2");

                entity.Property(e => e.Personnage).HasMaxLength(50);

                entity.Property(e => e.Texte).HasMaxLength(50);

                entity.Property(e => e.Type).HasMaxLength(50);
            });

            modelBuilder.Entity<Donsperso>(entity =>
            {
                entity.HasKey(e => e.Nom)
                    .HasName("PrimaryKey");

                entity.Property(e => e.Nom)
                    .HasMaxLength(50)
                    .ValueGeneratedNever();

                entity.Property(e => e.Caracteristique).HasMaxLength(50);

                entity.Property(e => e.Epique).HasColumnType("bit");

                entity.Property(e => e.Genre).HasMaxLength(255);

                entity.Property(e => e.Liearme).HasColumnType("bit");

                entity.Property(e => e.Multiple).HasColumnType("bit");

                entity.Property(e => e.Pouvoir).HasColumnType("bit");

                entity.Property(e => e.Prive).HasColumnType("bit");

                entity.Property(e => e.Source).HasMaxLength(50);

                entity.Property(e => e.Tare).HasColumnType("bit");

                entity.Property(e => e.Trait).HasColumnType("bit");
            });

            modelBuilder.Entity<Equipement>(entity =>
            {
                entity.HasIndex(e => e.NomObjet)
                    .HasName("EquipementnomObjet");

                entity.HasIndex(e => e.Personnage)
                    .HasName("{6C06414D-F300-491F-9149-676B98DC90D6}");

                entity.HasIndex(e => new { e.Personnage, e.Pos })
                    .HasName("index");

                entity.Property(e => e.NomObjet)
                    .HasColumnName("nomObjet")
                    .HasMaxLength(50);

                entity.Property(e => e.Personnage).HasMaxLength(50);

                entity.Property(e => e.Place)
                    .HasColumnName("place")
                    .HasDefaultValueSql("Null");

                entity.Property(e => e.Pos)
                    .HasColumnName("pos")
                    .HasDefaultValueSql("Null");

                entity.Property(e => e.SurPersonnage).HasColumnType("bit");

                entity.Property(e => e.Type).HasMaxLength(50);

                entity.HasOne(d => d.PersonnageNavigation)
                    .WithMany(p => p.Equipement)
                    .HasForeignKey(d => d.Personnage)
                    .HasConstraintName("{6C06414D-F300-491F-9149-676B98DC90D6}");
            });

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

            modelBuilder.Entity<ObjetsPropriete>(entity =>
            {
                entity.HasIndex(e => new { e.NomObjet, e.Propriete2 })
                    .HasName("index2");

                entity.HasIndex(e => new { e.NomObjet, e.Propriete1, e.Propriete2 })
                    .HasName("index1");

                entity.Property(e => e.NomObjet)
                    .HasColumnName("nomObjet")
                    .HasMaxLength(50);

                entity.Property(e => e.Propriete1)
                    .HasColumnName("Propriete_1")
                    .HasMaxLength(50);

                entity.Property(e => e.Propriete2)
                    .HasColumnName("Propriete_2")
                    .HasMaxLength(50);

                entity.Property(e => e.Propriete3)
                    .HasColumnName("Propriete_3")
                    .HasMaxLength(50);

                entity.Property(e => e.Valeur)
                    .HasColumnName("valeur")
                    .HasMaxLength(20);
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

            modelBuilder.Entity<PersonnageDons>(entity =>
            {
                entity.HasIndex(e => e.Nom)
                    .HasName("nom");

                entity.HasIndex(e => new { e.Nom, e.Classe, e.Niveau, e.Dons, e.Libelle })
                    .HasName("index")
                    .IsUnique();

                entity.Property(e => e.Classe)
                    .HasColumnName("classe")
                    .HasMaxLength(50);

                entity.Property(e => e.Dons).HasMaxLength(50);

                entity.Property(e => e.Libelle).HasMaxLength(50);

                entity.Property(e => e.Niv)
                    .HasColumnName("niv")
                    .HasDefaultValueSql("0");

                entity.Property(e => e.Niveau).HasDefaultValueSql("0");

                entity.Property(e => e.Nom)
                    .HasColumnName("nom")
                    .HasMaxLength(50);
            });

            modelBuilder.Entity<PersonnageProgression>(entity =>
            {
                entity.HasIndex(e => e.Nom)
                    .HasName("{4FCD1E6B-D6F2-454F-B043-167A177B5946}");

                entity.HasIndex(e => new { e.Nom, e.Niv })
                    .HasName("index1");

                entity.HasIndex(e => new { e.Nom, e.Classe, e.Niveau })
                    .HasName("index")
                    .IsUnique();

                entity.Property(e => e.Cha).HasDefaultValueSql("0");

                entity.Property(e => e.Classe)
                    .HasColumnName("classe")
                    .HasMaxLength(30);

                entity.Property(e => e.Competence)
                    .HasColumnName("competence")
                    .HasMaxLength(255);

                entity.Property(e => e.Con).HasDefaultValueSql("0");

                entity.Property(e => e.Dex).HasDefaultValueSql("0");

                entity.Property(e => e.For).HasDefaultValueSql("0");

                entity.Property(e => e.Int).HasDefaultValueSql("0");

                entity.Property(e => e.Modifint).HasColumnName("modifint");

                entity.Property(e => e.Niv)
                    .HasColumnName("niv")
                    .HasDefaultValueSql("0");

                entity.Property(e => e.Niveau).HasDefaultValueSql("0");

                entity.Property(e => e.Nom)
                    .HasColumnName("nom")
                    .HasMaxLength(50);

                entity.Property(e => e.Pv).HasDefaultValueSql("0");

                entity.Property(e => e.Sag).HasDefaultValueSql("0");

                entity.HasOne(d => d.NomNavigation)
                    .WithMany(p => p.PersonnageProgression)
                    .HasForeignKey(d => d.Nom)
                    .OnDelete(DeleteBehavior.Cascade)
                    .HasConstraintName("{4FCD1E6B-D6F2-454F-B043-167A177B5946}");
            });

            modelBuilder.Entity<SortListe>(entity =>
            {
                entity.Property(e => e.ClassePerso).HasMaxLength(255);

                entity.Property(e => e.EmplacementSort).HasMaxLength(50);

                entity.Property(e => e.Metamagie).HasMaxLength(255);

                entity.Property(e => e.NomPerso).HasMaxLength(255);

                entity.Property(e => e.NomSort).HasMaxLength(255);
            });

            modelBuilder.Entity<SortPersonnage>(entity =>
            {
                entity.HasIndex(e => e.NomPerso)
                    .HasName("PersonnageSortPersonnage");

                entity.HasIndex(e => new { e.NomPerso, e.ClassePerso, e.NomSort })
                    .HasName("index")
                    .IsUnique();

                entity.Property(e => e.ClassePerso).HasMaxLength(30);

                entity.Property(e => e.Niveau).HasMaxLength(4);

                entity.Property(e => e.NomPerso).HasMaxLength(50);

                entity.Property(e => e.NomSort).HasMaxLength(50);

                entity.HasOne(d => d.NomPersoNavigation)
                    .WithMany(p => p.SortPersonnage)
                    .HasForeignKey(d => d.NomPerso)
                    .OnDelete(DeleteBehavior.Cascade)
                    .HasConstraintName("PersonnageSortPersonnage");
            });

            modelBuilder.Entity<Version>(entity =>
            {
                entity.HasKey(e => e.Version1)
                    .HasName("PrimaryKey");

                entity.Property(e => e.Version1)
                    .HasColumnName("Version")
                    .ValueGeneratedNever();
            });
        }
    }
}
