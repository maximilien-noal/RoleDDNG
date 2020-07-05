using System;
using System.Collections.Generic;
using System.Text;

using GalaSoft.MvvmLight;

using RoleDDNG.ViewModels.Interfaces;

namespace RoleDDNG.ViewModels.ToolsVMs
{
    public class CharactersXpViewModel : ViewModelBase, IDocumentViewModel
    {
        public string Title => "Expérience des personnages";
    }
}