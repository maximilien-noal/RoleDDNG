using System;
using System.Collections.Generic;
using System.Text;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using RoleDDNG.ViewModels.Interfaces;

namespace RoleDDNG.ViewModels.ToolsVMs
{
    public class TownGeneratorViewModel : ViewModelBase, IContent
    {
        public TownGeneratorViewModel()
        {
        }

        public string Title => "Générateur de ville";
    }
}