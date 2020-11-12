using GalaSoft.MvvmLight.Command;

using System.Linq;

namespace RoleDDNG.ViewModels.Menus.Rules
{
    public class CollectionSlideShowViewModel<T> : CollectionViewModel<T>
        where T : new()
    {
        private T _selectedItem = new T();

        public CollectionSlideShowViewModel()
        {
            Previous = new RelayCommand(PreviousMethod);
            Next = new RelayCommand(NextMethod);
        }

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

        private void MoveTo(int diff)
        {
            if (SelectedItem is null && Collection.Any())
            {
                SelectedItem = Collection.First();
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
                SelectedItem = Collection.ElementAt(newIdex);
            }
        }
    }
}