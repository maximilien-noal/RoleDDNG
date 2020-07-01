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
        /// <summary>
        /// Identifies the <see cref="IsModal" /> dependency property.
        /// </summary>
        public static readonly DependencyProperty IsModalProperty =
            DependencyProperty.Register(nameof(IsModal), typeof(bool?), typeof(MdiContainer), new UIPropertyMetadata(OnIsModalChanged));

        static MdiContainer()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(MdiContainer), new FrameworkPropertyMetadata(typeof(MdiContainer)));
        }

        public MdiContainer() : base()
        {
            this.SelectionChanged += MdiContainer_SelectionChanged;
        }

        public bool? IsModal
        {
            get { return (bool?)GetValue(IsModalProperty); }

            set { SetValue(IsModalProperty, value); }
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

            if (newValue is IList newValueList)
            {
                InternalItemSource = newValueList;
            }
        }

        protected override void PrepareContainerForItemOverride(DependencyObject element, object item)
        {
            if (element is MdiWindow window)
            {
                window.FocusChanged += OnWindowFocusChanged;
                window.Closing += OnWindowClosing;
                window.WindowStateChanged += OnWindowStateChanged;
                window.Initialize(this);
                window.Position();
                window.Focus();
            }

            base.PrepareContainerForItemOverride(element, item);
        }

        private static void OnIsModalChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (e.NewValue == null || !((bool?)e.NewValue).HasValue) return;
            ((MdiContainer)d).SetCurrentValue(IsModalProperty, ((bool?)e.NewValue).Value);
        }

        private void MdiContainer_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0 && ItemContainerGenerator.ContainerFromItem(e.AddedItems[0]) is MdiWindow windowNew)
            {
                windowNew.SetValue(MdiWindow.IsSelectedProperty, true);
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

                // clear
                window.FocusChanged -= OnWindowFocusChanged;
                window.Closing -= OnWindowClosing;
                window.WindowStateChanged -= OnWindowStateChanged;
                window.DataContext = null;
            }
        }

        private void OnWindowFocusChanged(object sender, RoutedEventArgs e)
        {
            if (((MdiWindow)sender).IsFocused)
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
                SetCurrentValue(SelectedItemProperty, e.OriginalSource);

                ((MdiWindow)ItemContainerGenerator.ContainerFromItem(SelectedItem)).SetCurrentValue(MdiWindow.IsSelectedProperty, true);
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