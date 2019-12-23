using System.Reflection;
using System.Windows;
using PresentationTheme.Aero;

namespace RoleDDNG.Role
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            // Set theme resources
            AeroTheme.SetAsCurrentTheme();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            var splash = new SplashScreen(Assembly.GetExecutingAssembly(), "Assets/splashscreen.png");
            splash.Show(true, true);

            base.OnStartup(e);
        }
    }
}