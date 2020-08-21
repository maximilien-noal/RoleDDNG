using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using System.Collections.ObjectModel;
using System.Linq;

namespace RoleDDNG.ViewModels.Menus.Rules
{
    public class CollectionViewModel<T> : ViewModelBase
        where T : new()
    {
        private ObservableCollection<T> _collectionStore = new ObservableCollection<T>();

        private T _selectedItem = new T();

        public CollectionViewModel()
        {
            Previous = new RelayCommand(PreviousMethod);
            Next = new RelayCommand(NextMethod);
        }

        public ObservableCollection<T> Collection { get => _collectionStore; private set { Set(nameof(Collection), ref _collectionStore, value); } }

        public RelayCommand Next { get; private set; }

        public RelayCommand Previous { get; private set; }

        public T SelectedItem { get => _selectedItem; set { Set(nameof(SelectedItem), ref _selectedItem, value); } }

        protected void NextMethod()
        {
            MoveTo(1);
        }

        protected void PreviousMethod()
        {
            MoveTo(-1);
        }

        protected void SetCollection(ObservableCollection<T> newCollection)
        {
            Collection = newCollection;
        }

        private void MoveTo(int diff)
        {
            if (SelectedItem is null)
            {
                SelectedItem = Collection.FirstOrDefault();
                return;
            }
            var index = Collection.IndexOf(SelectedItem);
            if (index != -1)
            {
                var newIdex = index += diff;
                if (newIdex > Collection.Count - 1)
                {
                    newIdex = 0;
                }
                else if (newIdex < 0)
                {
                    newIdex = Collection.Count - 1;
                }
                SelectedItem = Collection.ElementAtOrDefault(newIdex);
            }
        }
    }
}