using GalaSoft.MvvmLight;
using System;
using System.Collections.Generic;
using System.Text;

namespace RoleDDNG.Models.Roles
{
    public class ChangeNomRace : ObservableObject
    {
        private string _oldRace = "";

        public string OldRace { get => _oldRace; set { Set(nameof(OldRace), ref _oldRace, value); } }

        private string _newRace = "";

        public string NewRace { get => _newRace; set { Set(nameof(NewRace), ref _newRace, value); } }
    }
}