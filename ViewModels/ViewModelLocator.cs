using GalaSoft.MvvmLight.Ioc;
using RoleDDNG.ViewModels.Menus.Rules;
using RoleDDNG.ViewModels.Menus.Tools;

namespace RoleDDNG.ViewModels
{
    /// <summary>
    /// This class contains static references to all the view models in the application and provides
    /// an entry point for the bindings.
    /// </summary>
    public class ViewModelLocator
    {
        public ViewModelLocator()
        {
            SimpleIoc.Default.Register(() => new MainViewModel());
        }

        // Justification : class must implement a parameterless public constructor for the XAML
        // side, where it initializes the SimpleIoc used here
#pragma warning disable CA1822

        public CharactersXpViewModel CharactersXp => new CharactersXpViewModel();

        public DiceRollViewModel DiceRoll => new DiceRollViewModel();

        public DonsDescriptionsViewModel DonsDescriptions => new DonsDescriptionsViewModel();

        public MainViewModel Main => SimpleIoc.Default.GetInstance<MainViewModel>();

        public OpenCharacterViewModel OpenCharacter => new OpenCharacterViewModel();

        public RacesDescriptionsViewModel RacesDescriptions => new RacesDescriptionsViewModel();

        public SpellsDescriptionsViewModel SpellsDescriptions => new SpellsDescriptionsViewModel();

        public TownGeneratorViewModel TownGenerator => new TownGeneratorViewModel();

#pragma warning restore CA1822
    }
}