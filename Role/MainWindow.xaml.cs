using RoleDDNG.Models.Structs;
using RoleDDNG.Role.Dialogs;
using RoleDDNG.Role.PInvoke;
using RoleDDNG.ViewModels;

using System.Windows;
using System.Windows.Interop;

namespace RoleDDNG.Role
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveAndExit();
        }

        private void AboutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            new AboutBox(this).ShowDialog();
        }

        private void ExitMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.Close();
        }

        private void SaveAndExit()
        {
            NativeMethods.GetWindowPlacement(new WindowInteropHelper(this).Handle, out WindowPlacement windowPlacement);
            ((MainViewModel)DataContext).AppSettings.MainWindowPlacement = windowPlacement;
            ((MainViewModel)DataContext).ExitApp.Execute(null);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var windowPlacement = ((MainViewModel)DataContext).AppSettings.MainWindowPlacement;
            NativeMethods.SetWindowPlacement(new WindowInteropHelper(this).Handle, ref windowPlacement);
        }
    }
}