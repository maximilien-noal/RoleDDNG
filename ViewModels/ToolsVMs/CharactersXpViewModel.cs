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
        private bool _isBusy = false;

        public bool IsBusy { get => _isBusy; set { Set(nameof(IsBusy), ref _isBusy, value); } }

        public async Task LoadCharactersDbDataAsync()
        {
            await Task.Delay(0).ConfigureAwait(false);
        }
    }
}