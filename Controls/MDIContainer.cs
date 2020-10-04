using System.Collections;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

using Hammer.MDI.Control.Events;

namespace Hammer.MDI.Control
{
    public sealed class MdiContainer : Selector
    {
        static MdiContainer()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(MdiContainer), new FrameworkPropertyMetadata(typeof(MdiContainer)));
        }

        public MdiContainer() : base()
        {
            this.SelectionChanged += MdiContainer_SelectionChanged;
        }

        internal int MinimizedWindowsCount { get; private set; }

        private IList InternalItemSource { get; set; } = new List<object>();

        protected override DependencyObject GetContainerForItemOverride()
        {
            return new MdiWindow();
        }

        protected override void OnItemsSourceChanged(IEnumerable oldValue, IEnumerable newValue)
        {
            base.OnItemsSourceChanged(oldValue, newValue);

            if (newValue is IList list)
            {
                InternalItemSource = list;
            }
        }

        protected override void PrepareContainerForItemOverride(DependencyObject element, object item)
        {
            if (element is MdiWindow window)
            {
                window.Loaded += OnWindowLoaded;
                window.PreviewMouseMove += Window_PreviewMouseMove;
                window.SizeChanged += OnWindowSizeChanged;
                window.FocusChanged += OnWindowFocusChanged;
                window.Closing += OnWindowClosing;
                window.WindowStateChanged += OnWindowStateChanged;
                window.Initialize(this);
                window.Position();
                window.Focus();
            }

            base.PrepareContainerForItemOverride(element, item);
        }

        private void Window_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (sender is MdiWindow window)
            {
                window.LockDimensionsOnce();
                window.UnlockDimensionsOnce();
            }
        }

        private void MdiContainer_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0 && ItemContainerGenerator.ContainerFromItem(e.AddedItems[0]) is MdiWindow window)
            {
                window.SetValue(MdiWindow.IsSelectedProperty, true);
                window.Focus();
            }
        }

        private void OnWindowClosing(object sender, RoutedEventArgs e)
        {
            var window = sender as MdiWindow;
            if (window?.DataContext != null)
            {
                InternalItemSource?.Remove(window.DataContext);
                if (Items.Count > 0)
                {
                    SetCurrentValue(SelectedItemProperty, Items[^1]);
                    if (ItemContainerGenerator.ContainerFromItem(SelectedItem) is MdiWindow windowNew)
                    {
                        windowNew.IsSelected = true;
                    }
                }
                window.Loaded -= OnWindowLoaded;
                window.PreviewMouseMove -= Window_PreviewMouseMove;
                window.FocusChanged -= OnWindowFocusChanged;
                window.Closing -= OnWindowClosing;
                window.WindowStateChanged -= OnWindowStateChanged;
                window.SizeChanged -= OnWindowSizeChanged;
                window.DataContext = null;
            }
        }

        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            if (sender is MdiWindow window)
            {
                window.FocusOnFirstFocusableChild();
            }
        }

        private void OnWindowFocusChanged(object sender, RoutedEventArgs e)
        {
            if (sender is MdiWindow senderWindow && senderWindow.IsFocused)
            {
                SetCurrentValue(SelectedItemProperty, e.OriginalSource);

                ((MdiWindow)ItemContainerGenerator.ContainerFromItem(SelectedItem)).SetCurrentValue(MdiWindow.IsSelectedProperty, true);

                foreach (var item in Items)
                {
                    if (item != e.OriginalSource && ItemContainerGenerator.ContainerFromItem(item) is MdiWindow window)
                    {
                        window.IsSelected = false;
                        Panel.SetZIndex(window, 0);
                    }
                }
            }
        }

        private void OnWindowSizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (sender is MdiWindow window && window.WindowState == WindowState.Normal)
            {
                window.ResizeToAvailableSpace();
                window.SaveFirstSize();
            }
        }

        private void OnWindowStateChanged(object sender, WindowStateChangedEventArgs e)
        {
            if (e.NewValue == WindowState.Minimized)
            {
                MinimizedWindowsCount++;
            }
            else if (e.OldValue == WindowState.Minimized)
            {
                MinimizedWindowsCount--;
            }
        }
    }
}