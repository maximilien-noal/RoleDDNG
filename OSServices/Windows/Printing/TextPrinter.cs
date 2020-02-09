using System;
using System.Drawing.Printing;
using System.Windows.Forms;

using RoleDDNG.Interfaces.Printing;

namespace RoleDDNG.OSServices.Windows.Printing
{
    public class TextPrinter : ITextPrinter
    {
        public void PrintText(string textToPrint)
        {
            if (string.IsNullOrWhiteSpace(textToPrint))
            {
                throw new ArgumentNullException(nameof(textToPrint));
            }
            using var docToPrint = new PrintDocument();

            using var printDialog = new PrintDialog
            {
                AllowSomePages = true,
                ShowHelp = true,
                AllowPrintToFile = true,
                ShowNetwork = true,
                AllowCurrentPage = true,
                Document = docToPrint
            };
            docToPrint.PrintPage += (s, e) =>
            {
                System.Drawing.Font printFont = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Regular);
                e.Graphics.DrawString(textToPrint, printFont, System.Drawing.Brushes.Black, 10, 10);
            };
            DialogResult result = printDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                docToPrint.Print();
            }
        }
    }
}