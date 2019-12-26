using System;
using System.Globalization;
using System.IO;
using System.Management;
using System.Reflection;
using System.Security.Principal;
using System.Windows;

using AdonisUI;

using Microsoft.Win32;

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
            WatchTheme();
            CHangeThemeIfWindowsChangedIt();
        }

        private Uri _currentTheme = default;

        private const string RegistryKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";

        private const string RegistryValueName = "AppsUseLightTheme";

        private void WatchTheme()
        {
            var currentUser = WindowsIdentity.GetCurrent();
            string query = string.Format(
                CultureInfo.InvariantCulture,
                @"SELECT * FROM RegistryValueChangeEvent WHERE Hive = 'HKEY_USERS' AND KeyPath = '{0}\\{1}' AND ValueName = '{2}'",
                currentUser.User.Value,
                RegistryKeyPath.Replace(@"\", @"\\", System.StringComparison.InvariantCulture),
                RegistryValueName);

            using var watcher = new ManagementEventWatcher(query);
            watcher.EventArrived += Watcher_EventArrived;

            // Start listening for events
            //watcher.Start(); //Disabled because InvalidOperationException is thrown by System.Uri on the third theme change at runtime...
        }

        private void Watcher_EventArrived(object sender, EventArrivedEventArgs e)
        {
            CHangeThemeIfWindowsChangedIt();
        }

        private void CHangeThemeIfWindowsChangedIt()
        {
            var newWindowsTheme = GetWindowsTheme();
            if (_currentTheme != newWindowsTheme)
            {
                _currentTheme = newWindowsTheme;
                ChangeTheme(_currentTheme);
            }
        }

        private static Uri GetWindowsTheme()
        {
            using RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath);
            object registryValueObject = key?.GetValue(RegistryValueName);
            if (registryValueObject == null)
            {
                return ResourceLocator.LightColorScheme;
            }

            int registryValue = (int)registryValueObject;

            return registryValue > 0 ? ResourceLocator.LightColorScheme : ResourceLocator.DarkColorScheme;
        }

        private static void ChangeTheme(Uri theme)
        {
            if (theme == ResourceLocator.DarkColorScheme)
            {
                ResourceLocator.SetColorScheme(Application.Current.Resources, ResourceLocator.DarkColorScheme, ResourceLocator.LightColorScheme);
            }
            else
            {
                ResourceLocator.SetColorScheme(Application.Current.Resources, ResourceLocator.LightColorScheme, ResourceLocator.DarkColorScheme);
            }
        }
    }
}