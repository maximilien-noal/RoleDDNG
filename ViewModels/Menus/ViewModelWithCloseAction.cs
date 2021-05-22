using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;

using RoleDDNG.ViewModels.Interfaces;

namespace RoleDDNG.ViewModels.Menus
{
    public class ViewModelWithCloseAction<T> : ViewModelBase where T : IDocumentViewModel
    {
        private bool _cancelPressed;

        public bool CancelPressed { get => _cancelPressed; set { Set(nameof(CancelPressed), ref _cancelPressed, value); } }

        public ViewModelWithCloseAction() => Cancel = new RelayCommand(() => CancelPressed = true);

        public RelayCommand Cancel { get; private set; }
    }
}