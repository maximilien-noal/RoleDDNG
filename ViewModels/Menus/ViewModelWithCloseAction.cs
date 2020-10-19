using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Ioc;

using RoleDDNG.ViewModels.Interfaces;

namespace RoleDDNG.ViewModels.Menus
{
    public class ViewModelWithCloseAction<T> : ViewModelBase where T : IDocumentViewModel
    {
        public ViewModelWithCloseAction()
        {
            Cancel = new RelayCommand(() => SimpleIoc.Default.GetInstance<MainViewModel>().RemoveDocumentViewModel<T>());
        }

        public RelayCommand Cancel { get; private set; }
    }
}