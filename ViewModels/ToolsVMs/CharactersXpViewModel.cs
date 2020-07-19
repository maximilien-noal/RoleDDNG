using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using GalaSoft.MvvmLight;

using RoleDDNG.ViewModels.Interfaces;

namespace RoleDDNG.ViewModels.ToolsVMs
{
    public class CharactersXpViewModel : ViewModelBase, IDocumentViewModel, ICharactersDbDependentViewModel
    {
        public string Title => "Expérience des personnages";

        public async Task LoadCharactersDbDataAsync()
        {
            await Task.Delay(0).ConfigureAwait(false);
        }
    }
}