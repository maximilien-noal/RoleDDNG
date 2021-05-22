namespace RoleDDNG.Role.UserControls
{
    using RoleDDNG.ViewModels.Interfaces;
    using RoleDDNG.ViewModels.Menus;

    using System;
    using System.Windows;

    /// <summary>
    /// Interaction logic for ChildWindow.xaml
    /// </summary>
    public partial class ChildWindow : Window
    {
        public ChildWindow()
        {
            InitializeComponent();
        }

        private void ChildWindowWin_SourceInitialized(object sender, EventArgs e)
        {
            if (DataContext is ViewModelWithCloseAction<IDocumentViewModel> vmc)
            {
                vmc.PropertyChanged += OnVmcPropertyChanged;
            }
        }

        private void OnVmcPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ViewModelWithCloseAction<IDocumentViewModel>.CancelPressed))
            {
                this.Close();
            }
        }
    }
}