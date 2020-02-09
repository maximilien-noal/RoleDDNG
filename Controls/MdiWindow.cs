using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;

using Hammer.MDI.Control.Events;
using Hammer.MDI.Control.Extensions;
using Hammer.MDI.Control.WindowControls;

namespace Hammer.MDI.Control
{
    [TemplatePart(Name = "PART_Border", Type = typeof(Border))]
    [TemplatePart(Name = "PART_BorderGrid", Type = typeof(Grid))]
    [TemplatePart(Name = "PART_Header", Type = typeof(Border))]
    [TemplatePart(Name = "PART_ButtonBar", Type = typeof(StackPanel))]
    [TemplatePart(Name = "PART_ButtonBar_MenuButton", Type = typeof(WindowButton))]
    [TemplatePart(Name = "PART_ButtonBar_CloseButton", Type = typeof(WindowButton))]
    [TemplatePart(Name = "PART_ButtonBar_MaximizeButton", Type = typeof(WindowButton))]
    [TemplatePart(Name = "PART_ButtonBar_MinimizeButton", Type = typeof(WindowButton))]
    [TemplatePart(Name = "PART_BorderContent", Type = typeof(Border))]
    [TemplatePart(Name = "PART_Content", Type = typeof(ContentPresenter))]
    [TemplatePart(Name = "PART_MoverThumb", Type = typeof(MoveThumb))]
    [TemplatePart(Name = "PART_ResizerThumb", Type = typeof(ResizeThumb))]
    [TemplatePart(Name = "PART_Thumblr", Type = typeof(Image))]
    [DebuggerDisplay("{Title}")]
    [TemplatePart(Name = "PART_Tumblr", Type = typeof(Image))]
    public sealed class MdiWindow : ContentControl
    {
        /// <summary> Identifies the <see cref="Closing" /> routed event. </summary>
        public static readonly RoutedEvent ClosingEvent = EventManager.RegisterRoutedEvent(
            nameof(Closing), RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(MdiWindow));

        /// <summary> Identifies the <see cref="FocusChanged" /> routed event. </summary>
        public static readonly RoutedEvent FocusChangedEvent = EventManager.RegisterRoutedEvent(
           nameof(FocusChanged), RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(MdiWindow));

        /// <summary> Identifies the <see cref="HasDropShadow" /> dependency property. </summary>
        public static readonly DependencyProperty HasDropShadowProperty =
            DependencyProperty.Register(nameof(HasDropShadow), typeof(bool), typeof(MdiWindow), new UIPropertyMetadata(true));

        /// <summary> Identifies the <see cref="IsModal" /> dependency property. </summary>
        public static readonly DependencyProperty IsModalProperty =
            DependencyProperty.Register(nameof(IsModal), typeof(bool?), typeof(MdiWindow), new UIPropertyMetadata(OnIsModalChanged));

        /// <summary> Identifies the <see cref="IsSelected" /> dependency property. </summary>
        public static readonly DependencyProperty IsSelectedProperty =
            DependencyProperty.Register(nameof(IsSelected), typeof(bool), typeof(MdiWindow), new UIPropertyMetadata(false));

        /// <summary> Identifies the <see cref="Title" /> dependency property. </summary>
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(nameof(Title), typeof(string), typeof(MdiWindow), new PropertyMetadata(string.Empty));

        /// <summary> Identifies the <see cref="WindowStateChanged" /> routed event. </summary>
        public static readonly RoutedEvent WindowStateChangedEvent = EventManager.RegisterRoutedEvent(
           nameof(WindowStateChanged), RoutingStrategy.Bubble, typeof(WindowStateChangedRoutedEventHandler), typeof(MdiWindow));

        /// <summary> Identifies the <see cref="WindowState" /> dependency property. </summary>
        public static readonly DependencyProperty WindowStateProperty =
            DependencyProperty.Register(nameof(WindowState), typeof(WindowState), typeof(MdiWindow), new PropertyMetadata(WindowState.Normal, OnWindowStateChanged));

        private WindowButton _closeButton;

        private WindowButton _maximizeButton;

        private WindowButton _menuButton;

        private WindowButton _minimizeButton;

        private Adorner _myAdorner;

        private AdornerLayer _myAdornerLayer;

        static MdiWindow()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(MdiWindow), new FrameworkPropertyMetadata(typeof(MdiWindow)));
        }

        public MdiWindow()
        {
            _myAdornerLayer = AdornerLayer.GetAdornerLayer(this);
        }

        public delegate void WindowStateChangedRoutedEventHandler(object sender, WindowStateChangedEventArgs e);

        public event RoutedEventHandler Closing
        {
            add { AddHandler(ClosingEvent, value); }
            remove { RemoveHandler(ClosingEvent, value); }
        }

        public event RoutedEventHandler FocusChanged
        {
            add { AddHandler(FocusChangedEvent, value); }
            remove { RemoveHandler(FocusChangedEvent, value); }
        }

        public event WindowStateChangedRoutedEventHandler WindowStateChanged
        {
            add { AddHandler(WindowStateChangedEvent, value); }
            remove { RemoveHandler(WindowStateChangedEvent, value); }
        }

        public bool HasDropShadow
        {
            get { return (bool)GetValue(HasDropShadowProperty); }
            set { SetValue(HasDropShadowProperty, value); }
        }

        public bool? IsModal
        {
            get
            {
                return (bool?)GetValue(IsModalProperty);
            }

            set
            {
#pragma warning disable WPF0036 // Avoid side effects in CLR accessors.
                if (!value.HasValue)
                {
                    _myAdornerLayer?.Remove(_myAdorner);
                }
                else
                {
                    if (_myAdornerLayer == null)
                        _myAdornerLayer = AdornerLayer.GetAdornerLayer(this);
                    if (_myAdorner == null)
                    {
                        _myAdorner = new HollowRectangleAdorner(this);
                    }
                    _myAdornerLayer.Add(_myAdorner);
                }
#pragma warning restore WPF0036 // Avoid side effects in CLR accessors.
                if (Container != null)
                {
                    Container.SetCurrentValue(MdiContainer.IsModalProperty, value);
                }

#pragma warning disable WPF0041 // Set mutable dependency properties using SetCurrentValue.
                SetValue(IsModalProperty, value);
#pragma warning restore WPF0041 // Set mutable dependency properties using SetCurrentValue.
            }
        }

        public bool IsSelected
        {
            get { return (bool)GetValue(IsSelectedProperty); }

            set { SetValue(IsSelectedProperty, value); }
        }

        public string Title
        {
            get { return (string)GetValue(TitleProperty); }
            set { SetValue(TitleProperty, value); }
        }

        public Image Tumblr { get; private set; }

        public WindowState WindowState
        {
            get
            {
                return (WindowState)GetValue(WindowStateProperty);
            }

            set
            {
                SetValue(WindowStateProperty, value);
            }
        }

        internal MdiContainer Container { get; private set; }

        internal double LastHeight { get; set; }

        internal double LastLeft { get; set; }

        internal double LastTop { get; set; }

        internal double LastWidth { get; set; }

        internal WindowState PreviousWindowState { get; set; }

        public void DoFocus(MouseButtonEventArgs mouseButtonEventArgs)
        {
            OnMouseLeftButtonDown(mouseButtonEventArgs);
        }

        public override void OnApplyTemplate()
        {
            _menuButton = GetTemplateChild("PART_ButtonBar_MenuButton") as WindowButton;
            if (_menuButton != null)
            {
                _menuButton.MouseDoubleClick += CloseWindow;
            }

            _closeButton = GetTemplateChild("PART_ButtonBar_CloseButton") as WindowButton;
            if (_closeButton != null)
            {
                _closeButton.Click += CloseWindow;
            }

            _maximizeButton = GetTemplateChild("PART_ButtonBar_MaximizeButton") as WindowButton;
            if (_maximizeButton != null)
            {
                _maximizeButton.Click += ToggleMaximizeWindow;
            }

            _minimizeButton = GetTemplateChild("PART_ButtonBar_MinimizeButton") as WindowButton;
            if (_minimizeButton != null)
            {
                _minimizeButton.Click += ToggleMinimizeWindow;
            }

            Tumblr = GetTemplateChild("PART_Tumblr") as Image;
        }

        public void Position()
        {
            var actualContainerHeight = Container.ActualHeight;
            var actualContainerWidth = Container.ActualWidth;
            UpdateLayout();
            InvalidateMeasure();
            var actualWidth = ActualWidth;
            var actualHeight = ActualHeight;

            var left = Math.Max(0, (actualContainerWidth - actualWidth) / 4);
            var top = Math.Max(0, (actualContainerHeight - actualHeight) / 4);

            SetCurrentValue(Canvas.LeftProperty, left);
            SetCurrentValue(Canvas.TopProperty, top);
        }

        internal void Initialize(MdiContainer container)
        {
            Container = container;
            Container.SizeChanged += OnContainerSizeChanged;
            LastHeight = ActualHeight;
            LastWidth = ActualWidth;
        }

        protected override void OnGotKeyboardFocus(KeyboardFocusChangedEventArgs e)
        {
            base.OnGotKeyboardFocus(e);

            SetCurrentValue(IsSelectedProperty, true);
            Panel.SetZIndex(this, 2);

            RaiseEvent(new RoutedEventArgs(FocusChangedEvent, DataContext));
        }

        protected override void OnLostKeyboardFocus(KeyboardFocusChangedEventArgs e)
        {
            if (e is null)
            {
                return;
            }

            base.OnLostKeyboardFocus(e);
            FrameworkElement parent = VisualTreeExtension.FindMdiWindow(e.NewFocus as FrameworkElement);
            if ((e.NewFocus is MdiWindow && !Equals(e.NewFocus, this)) || (parent != null && !Equals(parent, this)))
            {
                SetCurrentValue(IsSelectedProperty, false);
                Panel.SetZIndex(this, 0);
                var newWindow = (e.NewFocus is MdiWindow) ? (e.NewFocus as MdiWindow) : (parent as MdiWindow);
                Container.SetCurrentValue(MdiContainer.SelectedItemProperty, newWindow.DataContext);
                newWindow.IsSelected = true;
            }
        }

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            SetCurrentValue(IsSelectedProperty, true);
            Panel.SetZIndex(this, 2);
            RaiseEvent(new RoutedEventArgs(FocusChangedEvent, DataContext));

            Focus();
        }

        private static void OnIsModalChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (e.NewValue == null || ((bool?)e.NewValue).HasValue == false) return;
            ((MdiWindow)d).SetCurrentValue(IsModalProperty, ((bool?)e.NewValue).Value);
        }

        private static void OnWindowStateChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            if (obj is MdiWindow window)
            {
                window.PreviousWindowState = (WindowState)e.OldValue;

                var args = new WindowStateChangedEventArgs(WindowStateChangedEvent, (WindowState)e.OldValue, (WindowState)e.NewValue);
                window.RaiseEvent(args);
            }
        }

        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            RaiseEvent(new RoutedEventArgs(ClosingEvent));
        }

        private void OnContainerSizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (WindowState == WindowState.Maximized)
            {
                Width += e.NewSize.Width - e.PreviousSize.Width;
                Height += e.NewSize.Height - e.PreviousSize.Height;
            }
            if (WindowState == WindowState.Minimized)
            {
                Canvas.SetTop(this, Container.ActualHeight - 32);
            }
        }

        private void ToggleMaximizeWindow(object sender, RoutedEventArgs e)
        {
            Focus();
            this.ToggleMaximize();
        }

        private void ToggleMinimizeWindow(object sender, RoutedEventArgs e)
        {
            Focus();
            this.ToggleMinimize();
        }
    }
}