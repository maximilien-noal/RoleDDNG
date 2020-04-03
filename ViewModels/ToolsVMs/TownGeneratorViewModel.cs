using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

using AsyncAwaitBestPractices.MVVM;

using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.Interfaces.Printing;
using RoleDDNG.OSServices.CrossPlatform;
using RoleDDNG.ViewModels.Interfaces;

namespace RoleDDNG.ViewModels.ToolsVMs
{
    public class TownGeneratorViewModel : ViewModelBase, IContent
    {
        private readonly ITextPrinter _textPrinterService;

        private int _dimePercentage = 10;

        private bool _isBusy = false;

        private int _popCount = 500;

        private string _result = "";

        private int _taxPercentage = 10;

        private string _townName = "Nom";

        private string _townType = "Isolée";

        public TownGeneratorViewModel()
        {
            _textPrinterService = SimpleIoc.Default.GetInstance<ITextPrinter>();
            Generate = new AsyncCommand(GenerateMethodAsync);
            Print = new RelayCommand(PrintMethod);
        }

        public static List<string> TownTypes => new List<string>(new string[] { "Isolée", "Ouverte", "Intégrée" });

        public int DimePercentage { get => _dimePercentage; set { Set(nameof(DimePercentage), ref _dimePercentage, value); } }

        public AsyncCommand Generate { get; private set; }

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public int PopCount { get => _popCount; set { Set(nameof(PopCount), ref _popCount, value); } }

        public RelayCommand Print { get; private set; }

        public string Result { get => _result; set { Set(nameof(Result), ref _result, value); } }

        public int TaxPercentage { get => _taxPercentage; set { Set(nameof(TaxPercentage), ref _taxPercentage, value); } }

        public string Title => "Générateur de ville";

        public string TownName { get => _townName; set { Set(nameof(TownName), ref _townName, value); } }

        public string TownType { get => _townType; set { Set(nameof(TownType), ref _townType, value); } }

        private static int DiceRoll(int nombre, int face, int plus)
        {
            var dice = 0;
            for (int i = 0; i < nombre; i++)
            {
                dice += GetZeroIfNegative(Convert.ToInt32(face * new StaticRng().GetLimitedRNG().Next()) + 1);
            }
            dice += plus;
            return dice;
        }

        private static string GetAlignement()
        {
            var d100 = DiceRoll(1, 100, 0);
            if (d100 > 0)
            {
                return "Loyal bon";
            }
            if (d100 > 35)
            {
                return "Neutre bon";
            }
            if (d100 > 39)
            {
                return "Chaotique bon";
            }
            if (d100 > 41)
            {
                return "Loyal neutre";
            }
            if (d100 > 61)
            {
                return "Neutre";
            }
            if (d100 > 63)
            {
                return "Chaotique neutre";
            }
            if (d100 > 64)
            {
                return "Loyal mauvais";
            }
            if (d100 > 90)
            {
                return "Neutre mauvais";
            }
            if (d100 > 98)
            {
                return "Chaotique mauvais";
            }

            return "Chaotique bon";
        }

        private static int GetZeroIfNegative(int expr)
        {
            if (expr < 0)
            {
                return 0;
            }
            return expr;
        }

        private async Task GenerateMethodAsync()
        {
            IsBusy = true;
            Result = await Task.Run(() =>
            {
                var generatedTownInfo = new StringBuilder();
                generatedTownInfo.Append($"{TownName}, {GetTypeVille()} de {PopCount} habitants.{Environment.NewLine}");
                generatedTownInfo.Append($"Limite financière : {GetLimiteFinanciere()} Po{Environment.NewLine}");
                generatedTownInfo.Append($"Liquidite disponible : {GetLiquiditeDisponible()} Po{Environment.NewLine}");
                generatedTownInfo.Append($"{Environment.NewLine}Les revenus{Environment.NewLine}");
                generatedTownInfo.Append($"Revenus en or : {GetGoldRevenu()} Po/mois{Environment.NewLine}");
                generatedTownInfo.Append($"Revenus en matières premières : {GetRevenuFactor()} Po/mois{Environment.NewLine}");
                generatedTownInfo.Append($"Revenus totaux : {GetRevenu()} Po/mois{Environment.NewLine}");
                generatedTownInfo.Append($"{Environment.NewLine}Les impôts{Environment.NewLine}");
                generatedTownInfo.Append($"Impôts en or : {GetGoldRevenu() * TaxPercentage / 100} Po/mois{Environment.NewLine}");
                generatedTownInfo.Append($"Impôts en matières premières : {GetRevenuFactor() * TaxPercentage / 100} Po/mois{Environment.NewLine}");
                generatedTownInfo.Append($"Impôts totaux : {GetRevenu() * TaxPercentage / 100} Po/mois{Environment.NewLine}");
                generatedTownInfo.Append($"{Environment.NewLine}La dîme{Environment.NewLine}");
                generatedTownInfo.Append($"Dîme en or : {GetGoldRevenu() * DimePercentage / 100} Po/mois{Environment.NewLine}");
                generatedTownInfo.Append($"Dîme en matières premières : {GetRevenuFactor() * DimePercentage / 100} Po/mois{Environment.NewLine}");
                generatedTownInfo.Append($"Dîme totaux : {GetRevenu() * DimePercentage / 100} Po/mois{Environment.NewLine}");
                generatedTownInfo.Append($"{Environment.NewLine}Les instances{Environment.NewLine}");
                generatedTownInfo.Append($"{Instances()}{Environment.NewLine}");
                generatedTownInfo.Append($"{Environment.NewLine}Les PNJs{Environment.NewLine}");
                generatedTownInfo.Append($"{GetPNJs()}");
                generatedTownInfo.Append($"{Environment.NewLine}Mélange de races{Environment.NewLine}");
                generatedTownInfo.Append($"{GetRaces()}{Environment.NewLine}");
                return generatedTownInfo.ToString();
            }).ConfigureAwait(false);
            IsBusy = false;
        }

        private double GetGoldRevenu()
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

        private long GetLimiteFinanciere()
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

        private long GetLiquiditeDisponible()
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

        private string GetPNJs()
        {
            double commandant = Convert.ToInt32(new StaticRng().GetLimitedRNG().NextDouble() * 100) + 1;
            if (commandant < 61)
            {
                commandant = 8.1;
            }
            else if (commandant < 81)
            {
                commandant = 7.2;
            }
            else
            {
                commandant = 7.1;
            }

            var tabPnj = new Pnj[16];
            tabPnj[0].Classe = "Adepte"; tabPnj[0].Nombre = 1; tabPnj[0].Face = 6;
            tabPnj[1].Classe = "Barbare"; tabPnj[1].Nombre = 1; tabPnj[1].Face = 4;
            tabPnj[2].Classe = "Barde"; tabPnj[2].Nombre = 1; tabPnj[2].Face = 6;
            tabPnj[3].Classe = "Druide"; tabPnj[3].Nombre = 1; tabPnj[3].Face = 6;
            tabPnj[4].Classe = "Ensorceleur"; tabPnj[4].Nombre = 1; tabPnj[4].Face = 4;
            tabPnj[5].Classe = "Expert"; tabPnj[5].Nombre = 3; tabPnj[5].Face = 4;
            tabPnj[6].Classe = "Gens du peuple"; tabPnj[6].Nombre = 4; tabPnj[6].Face = 4;
            tabPnj[7].Classe = "Guerrier"; tabPnj[7].Nombre = 1; tabPnj[7].Face = 8;
            tabPnj[8].Classe = "Homme d'armes"; tabPnj[8].Nombre = 2; tabPnj[8].Face = 4;
            tabPnj[9].Classe = "Magicien"; tabPnj[9].Nombre = 1; tabPnj[9].Face = 4;
            tabPnj[10].Classe = "Moine"; tabPnj[10].Nombre = 1; tabPnj[10].Face = 4;
            tabPnj[11].Classe = "Noble"; tabPnj[11].Nombre = 1; tabPnj[12].Face = 4;
            tabPnj[12].Classe = "Paladin"; tabPnj[12].Nombre = 1; tabPnj[13].Face = 3;
            tabPnj[13].Classe = "Prêtre"; tabPnj[13].Nombre = 1; tabPnj[14].Face = 6;
            tabPnj[14].Classe = "Rôdeur"; tabPnj[14].Nombre = 1; tabPnj[15].Face = 3;
            tabPnj[15].Classe = "Roublard"; tabPnj[15].Nombre = 1; tabPnj[15].Face = 8;

            var modificateur = 0;
            if (PopCount >= 0)
            {
                modificateur = -3;
            }
            else if (PopCount > 80)
            {
                modificateur = -2;
            }
            else if (PopCount > 400)
            {
                modificateur = -1;
            }
            else if (PopCount > 900)
            {
                modificateur = 0;
            }
            else if (PopCount > 2000)
            {
                modificateur = 3;
            }
            else if (PopCount > 5000)
            {
                modificateur = 6;
            }
            else if (PopCount > 12000)
            {
                modificateur = 9;
            }
            else if (PopCount > 25000)
            {
                modificateur = 12;
            }
            var nombre = Math.Max(1, Convert.ToInt32(modificateur / 3));
            var totalPnj = 0;

            var niveauCommandant = 0;
            var pnjs = new StringBuilder();
            for (int j = 0; j < tabPnj.Length - 1; j++)
            {
                var bonusModificateur = 0;
                if ((j == 3 || j == 14) && modificateur < -1 && DiceRoll(1, 20, 0) == 20)
                {
                    bonusModificateur = 10;
                }
                var niveauPnj = new int[20];
                var niveau = 0;
                for (int i = 0; i < nombre; i++)
                {
                    niveau = DiceRoll(tabPnj[j].Nombre, tabPnj[j].Face, modificateur + bonusModificateur);

                    if (niveau > niveauPnj.Length - 1)
                    {
                        niveau = niveauPnj.Length - 1;
                    }
                    if (niveau > 0)
                    {
                        niveauPnj[niveau] = niveauPnj[niveau] + 1;
                        totalPnj += 1;
                        var multiplicateur = 1;
                        do
                        {
                            multiplicateur *= 2;
                            niveau = Convert.ToInt32((niveau + 1) / 2);
                            if (!(niveau == 1 &&
                                (tabPnj[j].Classe == "Adepte" ||
                                tabPnj[j].Classe == "Expert" ||
                                tabPnj[j].Classe == "Gens du peuple" ||
                                tabPnj[j].Classe == "Homme d'armes" ||
                                tabPnj[j].Classe == "Noble")))
                            {
                                niveauPnj[niveau] = niveauPnj[niveau] + multiplicateur;
                                totalPnj += multiplicateur;
                            }
                        } while (niveau > 1);
                    }
                    var k = 0;
                    if (niveau > 0)
                    {
                        pnjs.Append(tabPnj[j].Classe + " : ");
                        for (i = niveauPnj.Length - 1; i > 0; i--)
                        {
                            if (niveauPnj[i] != 0)
                            {
                                if (k != 0)
                                {
                                    pnjs.Append(", ");
                                }
                                else
                                {
                                    k++;
                                }
                                pnjs.Append($"{niveauPnj[i]} de niveau {i}");
                            }
                        }
                    }
                    k = -1;
                    if (j == Convert.ToInt32(commandant))
                    {
                        for (int l = niveauPnj.Length - 1; l > 0; l--)
                        {
                            if (niveauPnj[l] != 0)
                            {
                                if (k != -1 || (commandant - Convert.ToInt32(commandant)) * 10 < 1)
                                {
                                    niveauCommandant = l;
                                    break;
                                }
                                else
                                {
                                    k = 0;
                                }
                            }
                        }
                    }
                }
                if (niveau > 0)
                {
                    pnjs.Append(Environment.NewLine);
                }
            }

            double reste = PopCount - totalPnj;
            pnjs.Append($"{Environment.NewLine}Les classes de Pnj de niveau 1{Environment.NewLine}");
            pnjs.Append($"Gens du peuple : {GetZeroIfNegative(Convert.ToInt32(0.91 * reste))}{Environment.NewLine}");
            pnjs.Append($"Homme d'armes : {GetZeroIfNegative(Convert.ToInt32(0.05 * reste))}{Environment.NewLine}");
            pnjs.Append($"Expert : {GetZeroIfNegative(Convert.ToInt32(0.03 * reste))}{Environment.NewLine}");
            pnjs.Append($"Adepte : {GetZeroIfNegative(Convert.ToInt32(0.005 * reste))}{Environment.NewLine}");
            pnjs.Append($"Noble : {GetZeroIfNegative(Convert.ToInt32(0.005 * reste))}{Environment.NewLine}");

            pnjs.Append($"{Environment.NewLine}Commandant de la garde/prévôt{Environment.NewLine}");
            if (niveauCommandant > 0)
            {
                if (commandant == 8.1)
                {
                    pnjs.Append($"Homme d'armes du plus haut niveau {niveauCommandant}{Environment.NewLine}");
                }
                else if (commandant == 7.2)
                {
                    pnjs.Append($"Deuxième guerrier de la communauté {niveauCommandant}{Environment.NewLine}");
                }
                if (commandant == 7.1)
                {
                    pnjs.Append($"Guerrier de plus haut niveau {niveauCommandant}{Environment.NewLine}");
                }
            }
            else
            {
                pnjs.Append($"Il n'y a aucun guerrier ou homme d'armes dans cette communauté, choisir un autre PNJ{Environment.NewLine}");
            }
            return pnjs.ToString();
        }

        private string GetRaces()
        {
            var race = new StringBuilder();
            if (TownType == "Isolée")
            {
                race.Append($"Humains : {Convert.ToInt32(PopCount * 0.96)}{Environment.NewLine}");
                race.Append($"Halfelins : {Convert.ToInt32(PopCount * 0.02)}{Environment.NewLine}");
                race.Append($"Elfes : {Convert.ToInt32(PopCount * 0.1)}{Environment.NewLine}");
                race.Append($"Autres races : {Convert.ToInt32(PopCount * 0.01)}{Environment.NewLine}");
            }
            else if (TownType == "Ouverte")
            {
                race.Append($"Humains : {Convert.ToInt32(PopCount * 0.79)}{Environment.NewLine}");
                race.Append($"Halfelins : {Convert.ToInt32(PopCount * 0.09)}{Environment.NewLine}");
                race.Append($"Elfes : {Convert.ToInt32(PopCount * 0.05)}{Environment.NewLine}");
                race.Append($"Nains : {Convert.ToInt32(PopCount * 0.03)}{Environment.NewLine}");
                race.Append($"Gnomes : {Convert.ToInt32(PopCount * 0.02)}{Environment.NewLine}");
                race.Append($"Demi-elfes : {Convert.ToInt32(PopCount * 0.01)}{Environment.NewLine}");
                race.Append($"Demi-orques : {Convert.ToInt32(PopCount * 0.01)}{Environment.NewLine}");
            }
            else if (TownType == "Intégrée")
            {
                race.Append($"Humains : {Convert.ToInt32(PopCount * 0.37)}{Environment.NewLine}");
                race.Append($"Halfelins : {Convert.ToInt32(PopCount * 0.2)}{Environment.NewLine}");
                race.Append($"Elfes : {Convert.ToInt32(PopCount * 0.18)}{Environment.NewLine}");
                race.Append($"Nains : {Convert.ToInt32(PopCount * 0.1)}{Environment.NewLine}");
                race.Append($"Gnomes : {Convert.ToInt32(PopCount * 0.07)}{Environment.NewLine}");
                race.Append($"Demi-elfes : {Convert.ToInt32(PopCount * 0.05)}{Environment.NewLine}");
                race.Append($"Demi-orques : {Convert.ToInt32(PopCount * 0.03)}{Environment.NewLine}");
            }
            return race.ToString();
        }

        private double GetRevenu()
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

        private double GetRevenuFactor()
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

        private string GetTypeVille()
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

        private string Instances()
        {
            var number = 1;

            if (PopCount > 5000)
            {
                number += 1;
            }
            if (PopCount > 12000)
            {
                number += 1;
            }
            if (PopCount > 25000)
            {
                number += 1;
            }

            var d20 = 0;
            var instances = new StringBuilder();
            for (int i = 1; i <= number; i++)
            {
                if (instances.Length > 0)
                {
                    instances.Append(Environment.NewLine);
                }
                if (PopCount >= 0)
                {
                    d20 = DiceRoll(1, 20, -1);
                }
                if (PopCount > 80)
                {
                    d20 = DiceRoll(1, 20, 0);
                }
                if (PopCount > 400)
                {
                    d20 = DiceRoll(1, 20, 1);
                }
                if (PopCount > 900)
                {
                    d20 = DiceRoll(1, 20, 2);
                }
                if (PopCount > 2000)
                {
                    d20 = DiceRoll(1, 20, 3);
                }
                if (PopCount > 5000)
                {
                    d20 = DiceRoll(1, 20, 4);
                }
                if (PopCount > 12000)
                {
                    d20 = DiceRoll(1, 20, 5);
                }
                if (PopCount > 25000)
                {
                    d20 = DiceRoll(1, 20, 6);
                }
                if (d20 < 14)
                {
                    instances.Append("Traditionnel");
                    if (DiceRoll(1, 20, 0) == 20)
                    {
                        instances.Append(" et influence monstrueuse");
                    }
                }
                else
                {
                    if (d20 > 18)
                    {
                        instances.Append("Magique");
                    }
                    else
                    {
                        instances.Append("Inhabituel");
                    }
                }
                instances.Append($" d'alignement {GetAlignement()}");
            }

            return instances.ToString();
        }

        private void PrintMethod()
        {
            if (!string.IsNullOrWhiteSpace(Result))
            {
                _textPrinterService.PrintText(Result);
            }
        }

        private struct Pnj
        {
            public string Classe { get; set; }

            public int Face { get; set; }

            public int Nombre { get; set; }
        }
    }
}