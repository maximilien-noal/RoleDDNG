﻿using GalaSoft.MvvmLight;

using System;
using System.Collections.ObjectModel;
using System.Globalization;

namespace RoleDDNG.Models
{
    public class Personnage : ObservableObject
    {
        private string _nom = "";
        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }
        private string _race = "";
        public string Race { get => _race; set { Set(nameof(Race), ref _race, value); } }
        private string _classe1 = "";
        public string Classe1 { get => _classe1; set { Set(nameof(Classe1), ref _classe1, value); } }
        private short? _niv1 = 0;
        public short? Niv1 { get => _niv1; set { Set(nameof(Niv1), ref _niv1, value); } }
        private string _classe2 = "";
        public string Classe2 { get => _classe2; set { Set(nameof(Classe2), ref _classe2, value); } }
        private short? _niv2 = 0;
        public short? Niv2 { get => _niv2; set { Set(nameof(Niv2), ref _niv2, value); } }
        private string _classe3 = "";
        public string Classe3 { get => _classe3; set { Set(nameof(Classe3), ref _classe3, value); } }
        private short? _niv3 = 0;
        public short? Niv3 { get => _niv3; set { Set(nameof(Niv3), ref _niv3, value); } }
        private short? _endurance = 0;
        public short? Endurance { get => _endurance; set { Set(nameof(Endurance), ref _endurance, value); } }
        private short? _puissance = 0;
        public short? Puissance { get => _puissance; set { Set(nameof(Puissance), ref _puissance, value); } }
        private short? _précision = 0;
        public short? Précision { get => _précision; set { Set(nameof(Précision), ref _précision, value); } }
        private short? _équilibre = 0;
        public short? Équilibre { get => _équilibre; set { Set(nameof(Équilibre), ref _équilibre, value); } }
        private short? _résistance = 0;
        public short? Résistance { get => _résistance; set { Set(nameof(Résistance), ref _résistance, value); } }
        private short? _vitalité = 0;
        public short? Vitalité { get => _vitalité; set { Set(nameof(Vitalité), ref _vitalité, value); } }
        private short? _intellect = 0;
        public short? Intellect { get => _intellect; set { Set(nameof(Intellect), ref _intellect, value); } }
        private short? _érudition = 0;
        public short? Érudition { get => _érudition; set { Set(nameof(Érudition), ref _érudition, value); } }
        private short? _intuition = 0;
        public short? Intuition { get => _intuition; set { Set(nameof(Intuition), ref _intuition, value); } }
        private short? _volonté = 0;
        public short? Volonté { get => _volonté; set { Set(nameof(Volonté), ref _volonté, value); } }
        private short? _magnétisme = 0;
        public short? Magnétisme { get => _magnétisme; set { Set(nameof(Magnétisme), ref _magnétisme, value); } }
        private short? _charme = 0;
        public short? Charme { get => _charme; set { Set(nameof(Charme), ref _charme, value); } }
        private string _profil = "";
        public string Profil { get => _profil; set { Set(nameof(Profil), ref _profil, value); } }
        private string _titre = "";
        public string Titre { get => _titre; set { Set(nameof(Titre), ref _titre, value); } }
        private string _dieu = "";
        public string Dieu { get => _dieu; set { Set(nameof(Dieu), ref _dieu, value); } }
        private string _alignement = "";
        public string Alignement { get => _alignement; set { Set(nameof(Alignement), ref _alignement, value); } }
        private short? _beaute = 0;
        public short? Beaute { get => _beaute; set { Set(nameof(Beaute), ref _beaute, value); } }
        private string _cheveux = "";
        public string Cheveux { get => _cheveux; set { Set(nameof(Cheveux), ref _cheveux, value); } }
        private string _yeux = "";
        public string Yeux { get => _yeux; set { Set(nameof(Yeux), ref _yeux, value); } }
        private string _sexe = "";
        public string Sexe { get => _sexe; set { Set(nameof(Sexe), ref _sexe, value); } }
        private double? _poids = 0;
        public double? Poids { get => _poids; set { Set(nameof(Poids), ref _poids, value); } }
        private short? _taille = 0;
        public short? Taille { get => _taille; set { Set(nameof(Taille), ref _taille, value); } }
        private string _date = DateTime.MinValue.ToString(CultureInfo.GetCultureInfo("fr-FR"));
        public string Date { get => _date; set { Set(nameof(Date), ref _date, value); } }
        private short? _age = 0;
        public short? Age { get => _age; set { Set(nameof(Age), ref _age, value); } }
        private short? _maluxXp = 0;
        public short? MalusXp { get => _maluxXp; set { Set(nameof(MalusXp), ref _maluxXp, value); } }
        private string _ecoleSpecialisation = "";
        public string EcoleSpecialisation { get => _ecoleSpecialisation; set { Set(nameof(EcoleSpecialisation), ref _ecoleSpecialisation, value); } }
        private string _langueConnue = "";
        public string LangueConnue { get => _langueConnue; set { Set(nameof(LangueConnue), ref _langueConnue, value); } }
        private string _alphabet = "";
        public string Alphabet { get => _alphabet; set { Set(nameof(Alphabet), ref _alphabet, value); } }
        private bool _exclu = false;
        public bool Exclu { get => _exclu; set { Set(nameof(Exclu), ref _exclu, value); } }
        private string _domaine1 = "";
        public string Domaine1 { get => _domaine1; set { Set(nameof(Domaine1), ref _domaine1, value); } }
        private string _domaine2 = "";
        public string Domaine2 { get => _domaine2; set { Set(nameof(Domaine2), ref _domaine2, value); } }
        private string _domaine3 = "";
        public string Domaine3 { get => _domaine3; set { Set(nameof(Domaine3), ref _domaine3, value); } }
        private string _backGround = "";
        public string BackGround { get => _backGround; set { Set(nameof(BackGround), ref _backGround, value); } }
        private string _histoire = "";
        public string Histoire { get => _histoire; set { Set(nameof(Histoire), ref _histoire, value); } }
        private string _ecoleInterdite1 = "";
        public string EcoleInterdite1 { get => _ecoleInterdite1; set { Set(nameof(EcoleInterdite1), ref _ecoleInterdite1, value); } }
        private string _ecoleInterdite2 = "";
        public string EcoleInterdite2 { get => _ecoleInterdite2; set { Set(nameof(EcoleInterdite2), ref _ecoleInterdite2, value); } }
        private string _classe4 = "";
        public string Classe4 { get => _classe4; set { Set(nameof(Classe4), ref _classe4, value); } }
        private string _classe5 = "";
        public string Classe5 { get => _classe5; set { Set(nameof(Classe5), ref _classe5, value); } }
        private string _classe6 = "";
        public string Classe6 { get => _classe6; set { Set(nameof(Classe6), ref _classe6, value); } }
        private short? _niv4 = 0;
        public short? Niv4 { get => _niv4; set { Set(nameof(Niv4), ref _niv4, value); } }
        private short? _niv5 = 0;
        public short? Niv5 { get => _niv5; set { Set(nameof(Niv5), ref _niv5, value); } }
        private short? _niv6 = 0;
        public short? Niv6 { get => _niv6; set { Set(nameof(Niv6), ref _niv6, value); } }
        private string _archetype = "";
        public string Archetype { get => _archetype; set { Set(nameof(Archetype), ref _archetype, value); } }
        private double? _totalPo = 0d;
        public double? TotalPo { get => _totalPo; set { Set(nameof(TotalPo), ref _totalPo, value); } }
        private int? _totalXp = 0;
        public int? TotalXp { get => _totalXp; set { Set(nameof(TotalXp), ref _totalXp, value); } }
        private string _profession1 = "";
        public string Profession1 { get => _profession1; set { Set(nameof(Profession1), ref _profession1, value); } }
        private string _profession2 = "";
        public string Profession2 { get => _profession2; set { Set(nameof(Profession2), ref _profession2, value); } }
        private string _profession3 = "";
        public string Profession3 { get => _profession3; set { Set(nameof(Profession3), ref _profession3, value); } }
        private string _profession4 = "";
        public string Profession4 { get => _profession4; set { Set(nameof(Profession4), ref _profession4, value); } }
        private string _artisanat1 = "";
        public string Artisanat1 { get => _artisanat1; set { Set(nameof(Artisanat1), ref _artisanat1, value); } }
        private string _artisanat2 = "";
        public string Artisanat2 { get => _artisanat2; set { Set(nameof(Artisanat2), ref _artisanat2, value); } }
        private string _artisanat3 = "";
        public string Artisanat3 { get => _artisanat3; set { Set(nameof(Artisanat3), ref _artisanat3, value); } }
        private string _energieElu1 = "";
        public string EnergieElu1 { get => _energieElu1; set { Set(nameof(EnergieElu1), ref _energieElu1, value); } }
        private string _energieElu2 = "";
        public string EnergieElu2 { get => _energieElu2; set { Set(nameof(EnergieElu2), ref _energieElu2, value); } }
        private string _energieElu3 = "";
        public string EnergieElu3 { get => _energieElu3; set { Set(nameof(EnergieElu3), ref _energieElu3, value); } }
        private string _energieSorcier1 = "";
        public string EnergieSorcier1 { get => _energieSorcier1; set { Set(nameof(EnergieSorcier1), ref _energieSorcier1, value); } }
        private string _energieSorcier2 = "";
        public string EnergieSorcier2 { get => _energieSorcier2; set { Set(nameof(EnergieSorcier2), ref _energieSorcier2, value); } }
        private string _elementWujen = "";
        public string ElementWujen { get => _elementWujen; set { Set(nameof(ElementWujen), ref _elementWujen, value); } }
        private string _ennemisJures = "";
        public string EnnemisJures { get => _ennemisJures; set { Set(nameof(EnnemisJures), ref _ennemisJures, value); } }
        private string _image = "";
        public string Image { get => _image; set { Set(nameof(Image), ref _image, value); } }
        private string _domaine4 = "";
        public string Domaine4 { get => _domaine4; set { Set(nameof(Domaine4), ref _domaine4, value); } }
        private string _ecoleInterdite3 = "";
        public string EcoleInterdite3 { get => _ecoleInterdite3; set { Set(nameof(EcoleInterdite3), ref _ecoleInterdite3, value); } }
        private string _ecoleInterdite4 = "";
        public string EcoleInterdite4 { get => _ecoleInterdite4; set { Set(nameof(EcoleInterdite4), ref _ecoleInterdite4, value); } }
        private string _classe7 = "";
        public string Classe7 { get => _classe7; set { Set(nameof(Classe7), ref _classe7, value); } }
        private string _classe8 = "";
        public string Classe8 { get => _classe8; set { Set(nameof(Classe8), ref _classe8, value); } }
        private short? _niv7 = 0;
        public short? Niv7 { get => _niv7; set { Set(nameof(Niv7), ref _niv7, value); } }
        private short? _niv8 = 0;
        public short? Niv8 { get => _niv8; set { Set(nameof(Niv8), ref _niv8, value); } }
        private string _dragonTotem = "";
        public string DragonTotem { get => _dragonTotem; set { Set(nameof(DragonTotem), ref _dragonTotem, value); } }
        private string _elementShugenja = "";
        public string ElementShugenja { get => _elementShugenja; set { Set(nameof(ElementShugenja), ref _elementShugenja, value); } }
        private string _ordreShugenja = "";
        public string OrdreShugenja { get => _ordreShugenja; set { Set(nameof(OrdreShugenja), ref _ordreShugenja, value); } }

        public virtual ObservableCollection<Equipement> Equipement { get; private set; } = new ObservableCollection<Equipement>();
        public virtual ObservableCollection<PersonnageProgression> PersonnageProgression { get; private set; } = new ObservableCollection<PersonnageProgression>();
        public virtual ObservableCollection<SortPersonnage> SortPersonnage { get; private set; } = new ObservableCollection<SortPersonnage>();
    }
}