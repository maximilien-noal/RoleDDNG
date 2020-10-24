
using RoleDDNG.ViewModels.Interfaces;

using System.Collections.ObjectModel;

namespace RoleDDNG.ViewModels.Menus
{
    public class CollectionViewModel<T> : ViewModelWithCloseAction<IDocumentViewModel>
        where T : new()
    {
        private ObservableCollection<T> _collectionStore = new ObservableCollection<T>();

        public ObservableCollection<T> Collection { get => _collectionStore; private set { Set(nameof(Collection), ref _collectionStore, value); } }

        protected void SetCollection(ObservableCollection<T> newCollection)
        {
            Collection = newCollection;
        }
    }
}
