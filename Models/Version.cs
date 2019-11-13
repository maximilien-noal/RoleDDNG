using GalaSoft.MvvmLight;
using System;

namespace RoleDDNG.Models
{
    public class Version : ObservableObject
    {
        private short _version1 = 0;
        public short Version1 { get => _version1; set { Set(nameof(Version1), ref _version1, value); } }
        private DateTime? _dateModif = DateTime.MinValue;
        public DateTime? DateModif { get => _dateModif; set { Set(nameof(DateModif), ref _dateModif, value); } }
        private short? _objets = 0;
        public short? Objets { get => _objets; set { Set(nameof(Objets), ref _objets, value); } }
    }
}