using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;

namespace RoleDDNG.Role
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            var splash = new SplashScreen(Assembly.GetExecutingAssembly(), "Assets/splashscreen.png");
            splash.Show(true, true);

            base.OnStartup(e);
        }
    }
}