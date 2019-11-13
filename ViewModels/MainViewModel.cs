using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using System;
using System.Windows;

namespace RoleDDNG.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
#pragma warning disable CA1822 // Static bindings work, but makes the designer view throw an error.
        public RelayCommand ExitAppCommand { get => new RelayCommand(() => { Application.Current.Dispatcher.Invoke(() => Application.Current.MainWindow.Close()); }); }
    }

#pragma warning restore CA1822
}