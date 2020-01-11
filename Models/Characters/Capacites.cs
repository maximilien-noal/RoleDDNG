namespace RoleDDNG.Models.Characters
{
    using GalaSoft.MvvmLight;

    public class Capacites : ObservableObject
    {
        private string _nom = "";

        public string Nom { get => _nom; set { Set(nameof(Nom), ref _nom, value); } }

        private string _personnage = "";

        public string Personnage { get => _personnage; set { Set(nameof(Personnage), ref _personnage, value); } }

        private string _type = "";

        public string Type { get => _type; set { Set(nameof(Type), ref _type, value); } }

        private short? _modificateur1 = 0;

        public short? Modificateur1 { get => _modificateur1; set { Set(nameof(Modificateur1), ref _modificateur1, value); } }

        private short? _modificateur2 = 0;

        public short? Modificateur2 { get => _modificateur2; set { Set(nameof(Modificateur2), ref _modificateur2, value); } }

        private string _texte = "";

        public string Texte { get => _texte; set { Set(nameof(Texte), ref _texte, value); } }
    }
}