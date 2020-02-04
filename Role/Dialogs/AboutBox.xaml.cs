using System.Diagnostics;
using System.Windows;

namespace RoleDDNG.Role.Dialogs
{
    /// <summary>
    /// Interaction logic for AboutBox.xaml
    /// </summary>
    public partial class AboutBox : Window
    {
        public AboutBox(MainWindow mainWindow)
        {
            InitializeComponent();
            Owner = mainWindow;
        }

        public AboutBox()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

#pragma warning disable CC0091 // Use static method (bound from XAML side)

        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
#pragma warning restore CC0091 // Use static method (bound from XAML side)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }
    }
}