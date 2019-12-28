using System;
using System.Collections.Generic;
using System.Text;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;

using RoleDDNG.ViewModels.Interfaces;
using RoleDDNG.ViewModels.RNG;

namespace RoleDDNG.ViewModels.ToolsVMs
{
    public class TownGeneratorViewModel : ViewModelBase, IContent
    {
        public RelayCommand Generate { get; private set; }

        private string _townName = "";
        public string TownName { get => _townName; set { Set(nameof(TownName), ref _townName, value); } }

        private string _townType = "";

        public string TownType { get => _townType; set { Set(nameof(TownType), ref _townType, value); } }

        private int _popCount = 0;

        public int PopCount { get => _popCount; set { Set(nameof(PopCount), ref _popCount, value); } }

        private int _taxPercentage = 0;

        public int TaxPercentage { get => _taxPercentage; set { Set(nameof(TaxPercentage), ref _taxPercentage, value); } }

        private int _dimePercentage = 0;

        public int DimePercentage { get => _dimePercentage; set { Set(nameof(DimePercentage), ref _dimePercentage, value); } }

        private string _result = "";

        public string Result { get => _result; set { Set(nameof(Result), ref _result, value); } }

        public static List<string> TownTypes => new List<string>(new string[] { "Isolée", "Ouverte", "Intégrée" });
        public string Title => "Générateur de ville";

        public TownGeneratorViewModel()
        {
            Generate = new RelayCommand(GenerateMethod);
        }

        private void GenerateMethod()
        {
            var generatedTownInfo = new StringBuilder();
            generatedTownInfo.Append(TownName + ", " + TypeVille() + " de " + PopCount + " habitants." + Environment.NewLine);
            generatedTownInfo.Append("Limite financière : " + LimiteFinanciere() + " Po" + Environment.NewLine);
            generatedTownInfo.Append("Liquidite disponible : " + LiquiditeDisponible() + " Po" + Environment.NewLine);
            generatedTownInfo.Append(Environment.NewLine + "Les revenus" + Environment.NewLine);
            generatedTownInfo.Append("Revenus en or : " + RevenuOr() + " Po/mois" + Environment.NewLine);
            generatedTownInfo.Append("Revenus en matières premières : " + RevenuPremiere() + " Po/mois" + Environment.NewLine);
            generatedTownInfo.Append("Revenus totaux : " + Revenu() + " Po/mois" + Environment.NewLine);
            generatedTownInfo.Append(Environment.NewLine + "Les impôts" + Environment.NewLine);
            generatedTownInfo.Append("Impôts en or : " + RevenuOr() * TaxPercentage / 100 + " Po/mois" + Environment.NewLine);
            generatedTownInfo.Append("Impôts en matières premières : " + RevenuPremiere() * TaxPercentage / 100 + " Po/mois" + Environment.NewLine);
            generatedTownInfo.Append("Impôts totaux : " + Revenu() * TaxPercentage / 100 + " Po/mois" + Environment.NewLine);
            generatedTownInfo.Append(Environment.NewLine + "La dîme" + Environment.NewLine);
            generatedTownInfo.Append("Dîme en or : " + RevenuOr() * DimePercentage / 100 + " Po/mois" + Environment.NewLine);
            generatedTownInfo.Append("Dîme en matières premières : " + RevenuPremiere() * DimePercentage / 100 + " Po/mois" + Environment.NewLine);
            generatedTownInfo.Append("Dîme totaux : " + Revenu() * DimePercentage / 100 + " Po/mois" + Environment.NewLine);
            generatedTownInfo.Append(Environment.NewLine + "Les instances" + Environment.NewLine);
            generatedTownInfo.Append(Instances() + Environment.NewLine);
            generatedTownInfo.Append(Environment.NewLine + "Les Pnjs" + Environment.NewLine);
            //fldstrSortie.Append(Pnj(PopCount); //TODO : Translate PNJ method
            generatedTownInfo.Append(Environment.NewLine + "Mélange de races" + Environment.NewLine);
            generatedTownInfo.Append(Race() + Environment.NewLine);
            Result = generatedTownInfo.ToString();
        }

        private string Instances()
        {
            var Nombre = 1;

            if (PopCount > 5000)
            {
                Nombre = Nombre + 1;
            }
            if (PopCount > 12000)
            {
                Nombre = Nombre + 1;
            }
            if (PopCount > 25000)
            {
                Nombre = Nombre + 1;
            }

            var D20 = 0;
            var instances = "";
            for (int i = 0; i < Nombre; i++)
            {
                if (string.IsNullOrWhiteSpace(instances) == false)
                {
                    instances = instances + ", ";
                }
                if (PopCount >= 0)
                {
                    D20 = Des(1, 20, -1);
                }
                if (PopCount > 80)
                {
                    D20 = Des(1, 20, 0);
                }
                if (PopCount > 400)
                {
                    D20 = Des(1, 20, 1);
                }
                if (PopCount > 900)
                {
                    D20 = Des(1, 20, 2);
                }
                if (PopCount > 2000)
                {
                    D20 = Des(1, 20, 3);
                }
                if (PopCount > 5000)
                {
                    D20 = Des(1, 20, 4);
                }
                if (PopCount > 12000)
                {
                    D20 = Des(1, 20, 5);
                }
                if (PopCount > 25000)
                {
                    D20 = Des(1, 20, 6);
                }
                if (D20 < 14)
                {
                    instances = instances + "Traditionnel";
                    if (Des(1, 20, 0) == 20)
                    {
                        instances = instances + " et influence monstrueuse";
                    }
                    else
                    {
                        if (D20 > 18)
                        {
                            instances = instances + "Magique";
                        }
                        else
                        {
                            instances = instances + "Inhabituel";
                        }
                    }
                }
                instances = instances + " d'alignement " + Alignement();
            }

            return instances;
        }

        private static int Des(int nombre, int face, int plus)
        {
            var des = 0;
            for (int i = 0; i < nombre; i++)
            {
                des = des + Convert.ToInt32(face * StaticRNG.RNG.Next()) + 1;
            }
            des = des + plus;
            return des;
        }

        private static string Alignement()
        {
            var D100 = 0;

            D100 = Des(1, 100, 0);

            var alignement = "";

            if (D100 > 0)
            {
                alignement = "Loyal bon";
            }
            if (D100 > 35)
            {
                alignement = "Neutre bon";
            }
            if (D100 > 39)
            {
                alignement = "Chaotique bon";
            }
            if (D100 > 41)
            {
                alignement = "Loyal neutre";
            }
            if (D100 > 61)
            {
                alignement = "Neutre";
            }
            if (D100 > 63)
            {
                alignement = "Chaotique neutre";
            }
            if (D100 > 64)
            {
                alignement = "Loyal mauvais";
            }
            if (D100 > 90)
            {
                alignement = "Neutre mauvais";
            }
            if (D100 > 98)
            {
                alignement = "Chaotique mauvais";
            }

            return alignement;
        }

        private string Race()
        {
            var race = "";
            if (TownType == "Isolée")
            {
                race = race + "Humains : " + Convert.ToInt32(PopCount * 0.96) + Environment.NewLine;
                race = race + "Halfelins : " + Convert.ToInt32(PopCount * 0.02) + Environment.NewLine;
                race = race + "Elfes : " + Convert.ToInt32(PopCount * 0.1) + Environment.NewLine;
                race = race + "Autres races : " + Convert.ToInt32(PopCount * 0.01) + Environment.NewLine;
            }
            else if (TownType == "Ouverte")
            {
                race = race + "Humains : " + Convert.ToInt32(PopCount * 0.79) + Environment.NewLine;
                race = race + "Halfelins : " + Convert.ToInt32(PopCount * 0.09) + Environment.NewLine;
                race = race + "Elfes : " + Convert.ToInt32(PopCount * 0.05) + Environment.NewLine;
                race = race + "Nains : " + Convert.ToInt32(PopCount * 0.03) + Environment.NewLine;
                race = race + "Gnomes : " + Convert.ToInt32(PopCount * 0.02) + Environment.NewLine;
                race = race + "Demi-elfes : " + Convert.ToInt32(PopCount * 0.01) + Environment.NewLine;
                race = race + "Demi-orques : " + Convert.ToInt32(PopCount * 0.01) + Environment.NewLine;
            }
            else if (TownType == "Intégrée")
            {
                race = race + "Humains : " + Convert.ToInt32(PopCount * 0.37) + Environment.NewLine;
                race = race + "Halfelins : " + Convert.ToInt32(PopCount * 0.2) + Environment.NewLine;
                race = race + "Elfes : " + Convert.ToInt32(PopCount * 0.18) + Environment.NewLine;
                race = race + "Nains : " + Convert.ToInt32(PopCount * 0.1) + Environment.NewLine;
                race = race + "Gnomes : " + Convert.ToInt32(PopCount * 0.07) + Environment.NewLine;
                race = race + "Demi-elfes : " + Convert.ToInt32(PopCount * 0.05) + Environment.NewLine;
                race = race + "Demi-orques : " + Convert.ToInt32(PopCount * 0.03) + Environment.NewLine;
            }
            return race;
        }

        private double RevenuPremiere()
        {
            if (PopCount <= 80)
            {
                return 2.5 * PopCount * 0.95;
            }
            if (PopCount <= 400)
            {
                return 3 * PopCount * 0.9;
            }
            if (PopCount <= 900)
            {
                return 3.7 * PopCount * 0.8;
            }
            if (PopCount <= 2000)
            {
                return 4.5 * PopCount * 0.7;
            }
            if (PopCount <= 5000)
            {
                return 5.5 * PopCount * 0.6;
            }
            if (PopCount <= 12000)
            {
                return 6.7 * PopCount * 0.5;
            }
            if (PopCount <= 25000)
            {
                return 8.2 * PopCount * 0.35;
            }
            return 10 * PopCount * 0.2;
        }

        private double RevenuOr()
        {
            if (PopCount <= 80)
            {
                return 2.5 * PopCount * 0.05;
            }
            if (PopCount <= 400)
            {
                return 3 * PopCount * 0.1;
            }
            if (PopCount <= 900)
            {
                return 3.7 * PopCount * 0.2;
            }
            if (PopCount <= 2000)
            {
                return 4.5 * PopCount * 0.3;
            }
            if (PopCount <= 5000)
            {
                return 5.5 * PopCount * 0.4;
            }
            if (PopCount <= 12000)
            {
                return 6.7 * PopCount * 0.5;
            }
            if (PopCount <= 25000)
            {
                return 8.2 * PopCount * 0.65;
            }
            return 10 * PopCount * 0.8;
        }

        private double Revenu()
        {
            if (PopCount <= 80)
            {
                return 2.5 * PopCount;
            }
            if (PopCount <= 400)
            {
                return 3 * PopCount;
            }
            if (PopCount <= 900)
            {
                return 3.7 * PopCount;
            }
            if (PopCount <= 2000)
            {
                return 4.5 * PopCount;
            }
            if (PopCount <= 5000)
            {
                return 5.5 * PopCount;
            }
            if (PopCount <= 12000)
            {
                return 6.7 * PopCount;
            }
            if (PopCount <= 25000)
            {
                return 8.2 * PopCount;
            }
            return 10 * PopCount;
        }

        private long LiquiditeDisponible()
        {
            if (PopCount <= 80)
            {
                return 2 * PopCount;
            }
            if (PopCount <= 400)
            {
                return 5 * PopCount;
            }
            if (PopCount <= 900)
            {
                return 10 * PopCount;
            }
            if (PopCount <= 2000)
            {
                return 40 * PopCount;
            }
            if (PopCount <= 5000)
            {
                return 150 * PopCount;
            }
            if (PopCount <= 12000)
            {
                return 750 * PopCount;
            }
            if (PopCount <= 25000)
            {
                return 2000 * PopCount;
            }
            return 5000 * PopCount;
        }

        private long LimiteFinanciere()
        {
            if (PopCount <= 0)
            {
                return 0;
            }
            if (PopCount <= 80)
            {
                return 40;
            }
            if (PopCount <= 400)
            {
                return 100;
            }
            if (PopCount <= 900)
            {
                return 200;
            }
            if (PopCount <= 2000)
            {
                return 800;
            }
            if (PopCount <= 5000)
            {
                return 3000;
            }
            if (PopCount <= 12000)
            {
                return 15000;
            }
            if (PopCount <= 25000)
            {
                return 40000;
            }
            return 100000;
        }

        private string TypeVille()
        {
            if (PopCount <= 80)
            {
                return "Lieu dit";
            }
            if (PopCount <= 400)
            {
                return "Hameau";
            }
            if (PopCount <= 900)
            {
                return "Village";
            }
            if (PopCount <= 2000)
            {
                return "Bourg";
            }
            if (PopCount <= 5000)
            {
                return "Ville";
            }
            if (PopCount <= 12000)
            {
                return "Grande ville";
            }
            if (PopCount <= 25000)
            {
                return "Cité";
            }
            return "Métropole";
        }
    }
}