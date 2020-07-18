/*
  In App.xaml:
  <Application.Resources>
      <vm:ViewModelLocator xmlns:vm="clr-namespace:LTFGameLauncher"
                           x:Key="Locator" />
  </Application.Resources>

  In the View:
  DataContext="{Binding Source={StaticResource Locator}, Path=ViewModelName}"

  You can also use Blend to do all this with the tool's support.
  See http://www.galasoft.ch/mvvm
*/

using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.ViewModels.ToolsVMs;

namespace RoleDDNG.ViewModels
{
    /// <summary>
    /// This class contains static references to all the view models in the application and provides
    /// an entry point for the bindings.
    /// </summary>
    public class ViewModelLocator
    {
        /// <summary>
        /// Initializes a new instance of the ViewModelLocator class.
        /// </summary>
        public ViewModelLocator()
        {
            ////if (ViewModelBase.IsInDesignModeStatic)
            ////{
            ////    // Create design time view services and models
            ////    SimpleIoc.Default.Register<IDataService, DesignDataService>();
            ////}
            ////else
            ////{
            ////    // Create run time view services and models
            ////    SimpleIoc.Default.Register<IDataService, DataService>();
            ////}

            SimpleIoc.Default.Register(() => new MainViewModel());
            SimpleIoc.Default.Register<DiceRollViewModel>();
            SimpleIoc.Default.Register<TownGeneratorViewModel>();
            SimpleIoc.Default.Register<OpenCharacterViewModel>();
        }

#pragma warning disable CA1822

        // Justification : class must implement a parameterless public constructor for the XAML
        // side, where it initializes the SimpleIoc used here
        public MainViewModel Main
        {
            get
            {
                return SimpleIoc.Default.GetInstance<MainViewModel>();
            }
        }

        public CharactersXpViewModel CharactersXp
        {
            get
            {
                return SimpleIoc.Default.GetInstance<CharactersXpViewModel>();
            }
        }

        public DiceRollViewModel DiceRoll
        {
            get
            {
                return SimpleIoc.Default.GetInstance<DiceRollViewModel>();
            }
        }

        public TownGeneratorViewModel TownGenerator
        {
            get
            {
                return SimpleIoc.Default.GetInstance<TownGeneratorViewModel>();
            }
        }

        public OpenCharacterViewModel OpenCharacter
        {
            get
            {
                return SimpleIoc.Default.GetInstance<OpenCharacterViewModel>();
            }
        }

#pragma warning restore CA1822

        internal static void Cleanup()
        {
            //Nothing to do here.
        }
    }
}