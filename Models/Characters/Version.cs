using System;

using GalaSoft.MvvmLight;

namespace RoleDDNG.Models.Characters
{
    public class Version : ObservableObject
    {
        private DateTime? _dateModif = DateTime.MinValue;

        private short? _objets;

        private short _version1;

        public Version()
        {
        }

        public DateTime? DateModif { get => _dateModif; set { Set(nameof(DateModif), ref _dateModif, value); } }

        public short? Objets { get => _objets; set { Set(nameof(Objets), ref _objets, value); } }

        public short Version1 { get => _version1; set { Set(nameof(Version1), ref _version1, value); } }
    }
}