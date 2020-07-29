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
            SimpleIoc.Default.Register<DiceRollViewModel>();
            SimpleIoc.Default.Register<TownGeneratorViewModel>();
            SimpleIoc.Default.Register<OpenCharacterViewModel>();
            SimpleIoc.Default.Register<CharactersXpViewModel>();
            SimpleIoc.Default.Register<RacesDescriptionsViewModel>();
        }

        // Justification : class must implement a parameterless public constructor for the XAML
        // side, where it initializes the SimpleIoc used here
#pragma warning disable CA1822

        public CharactersXpViewModel CharactersXp => SimpleIoc.Default.GetInstance<CharactersXpViewModel>();

        public DiceRollViewModel DiceRoll => SimpleIoc.Default.GetInstance<DiceRollViewModel>();

        public MainViewModel Main => SimpleIoc.Default.GetInstance<MainViewModel>();

        public OpenCharacterViewModel OpenCharacter => SimpleIoc.Default.GetInstance<OpenCharacterViewModel>();

        public RacesDescriptionsViewModel RacesDescriptions => SimpleIoc.Default.GetInstance<RacesDescriptionsViewModel>();

        public TownGeneratorViewModel TownGenerator => SimpleIoc.Default.GetInstance<TownGeneratorViewModel>();

#pragma warning restore CA1822
    }
}