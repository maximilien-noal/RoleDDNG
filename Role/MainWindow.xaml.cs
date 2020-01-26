using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

using GalaSoft.MvvmLight.Command;

using Microsoft.VisualStudio.Threading;

using RoleDDNG.Models.Structs;
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
        private bool _forceClose = false;

        public RelayCommand HelpCommand { get; private set; }

        public MainWindow()
        {
            HelpCommand = new RelayCommand(HelpCommandMethod);
            InitializeComponent();
            Closing += MainWindow_Closing;
        }

        private void HelpCommandMethod()
        {
            AboutMenuItem_Click(this, new RoutedEventArgs());
        }

#pragma warning disable VSTHRD100 // Avoid async void methods (this is an event)

        private async void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
#pragma warning restore VSTHRD100 // Avoid async void methods (this is an event)
        {
            e.Cancel = true;
            if (_forceClose)
            {
                e.Cancel = false;
                return;
            }
            await SaveAndExitAsync().ConfigureAwait(true);
        }

        public void CloseForced()
        {
            _forceClose = true;
            Close();
        }

        private void AboutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            new AboutBox(this).ShowDialog();
        }

        private void ExitMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.Close();
        }

        private async Task SaveAndExitAsync()
        {
            NativeMethods.GetWindowPlacement(new WindowInteropHelper(this).Handle, out WindowPlacement windowPlacement);
            ((MainViewModel)DataContext).AppSettings.MainWindowPlacement = windowPlacement;
            await ((MainViewModel)DataContext).ExitApp.ExecuteAsync().ConfigureAwait(true);
            CloseForced();
        }

        private void Window_SourceInitialized(object sender, System.EventArgs e)
        {
            var windowPlacement = ((MainViewModel)DataContext).AppSettings.MainWindowPlacement;
            NativeMethods.SetWindowPlacement(new WindowInteropHelper(this).Handle, ref windowPlacement);
        }
    }
}