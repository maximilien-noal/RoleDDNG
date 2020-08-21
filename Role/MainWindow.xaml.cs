using System.Windows;
using System.Windows.Interop;

using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.Serialization;
using RoleDDNG.Interfaces.Backgrounds;
using RoleDDNG.Interfaces.Dialogs;
using RoleDDNG.Interfaces.Printing;
using RoleDDNG.Interfaces.Serialization;
using RoleDDNG.Interfaces.Window;
using RoleDDNG.Models.Options;
using RoleDDNG.Models.Structs;
using RoleDDNG.OSServices.Windows.Dialogs;
using RoleDDNG.OSServices.Windows.Printing;
using RoleDDNG.Role.Backgrounds;
using RoleDDNG.Role.Dialogs;
using RoleDDNG.Role.PInvoke;
using RoleDDNG.ViewModels;

namespace RoleDDNG.Role
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            SimpleIoc.Default.Register<IWindowPlacer, HwndPlacer>();
            SimpleIoc.Default.Register<IBackgroundSource, BackgroundSource>();
            SimpleIoc.Default.Register<IAsyncSerializer<AppSettings>, ModelSerializer<AppSettings>>();
            SimpleIoc.Default.Register<ITextPrinter, TextPrinter>();
            SimpleIoc.Default.Register<IFileDialog, Win32FileDialog>();
            InitializeComponent();
            HelpCommand = new RelayCommand(HelpCommandMethod);
            Closing += MainWindow_Closing;
        }

        public RelayCommand HelpCommand { get; private set; }

        private void AboutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            new AboutBox(this).ShowDialog();
        }

        private void ExitMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.Close();
        }

        private void HelpCommandMethod()
        {
            AboutMenuItem_Click(this, new RoutedEventArgs());
        }

#pragma warning disable VSTHRD100 // Avoid async void methods (this is an event)

        private async void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Closing -= MainWindow_Closing;
            SimpleIoc.Default.GetInstance<IWindowPlacer>().GetWindowPlacement(new WindowInteropHelper(this).Handle, out WindowPlacement windowPlacement);
            SimpleIoc.Default.GetInstance<AppSettings>().MainWindowPlacement = windowPlacement;
            await SimpleIoc.Default.GetInstance<MainViewModel>().ExitAppAsync().ConfigureAwait(true);
        }

        private async void Window_SourceInitialized(object sender, System.EventArgs e)
        {
            var foundCharacterDb = await SimpleIoc.Default.GetInstance<MainViewModel>().LoadAppSettingsAsync().ConfigureAwait(true);
            var windowPlacement = SimpleIoc.Default.GetInstance<AppSettings>().MainWindowPlacement;
            SimpleIoc.Default.GetInstance<IWindowPlacer>().SetWindowPlacement(new WindowInteropHelper(this).Handle, ref windowPlacement);
            if (foundCharacterDb)
            {
                await SimpleIoc.Default.GetInstance<MainViewModel>().OpenCharacterSheet.ExecuteAsync().ConfigureAwait(true);
            }
        }

#pragma warning restore VSTHRD100 // Avoid async void methods (this is an event)
    }
}