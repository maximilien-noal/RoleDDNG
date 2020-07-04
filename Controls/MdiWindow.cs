using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

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
        /// <summary>
        /// Identifies the <see cref="Closing" /> routed event.
        /// </summary>
        public static readonly RoutedEvent ClosingEvent = EventManager.RegisterRoutedEvent(
            nameof(Closing), RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(MdiWindow));

        /// <summary>
        /// Identifies the <see cref="FocusChanged" /> routed event.
        /// </summary>
        public static readonly RoutedEvent FocusChangedEvent = EventManager.RegisterRoutedEvent(
           nameof(FocusChanged), RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(MdiWindow));

        /// <summary>
        /// Identifies the <see cref="HasDropShadow" /> dependency property.
        /// </summary>
        public static readonly DependencyProperty HasDropShadowProperty =
            DependencyProperty.Register(nameof(HasDropShadow), typeof(bool), typeof(MdiWindow), new UIPropertyMetadata(true));

        /// <summary>
        /// Identifies the <see cref="IsModal" /> dependency property.
        /// </summary>
        public static readonly DependencyProperty IsModalProperty =
            DependencyProperty.Register(nameof(IsModal), typeof(bool?), typeof(MdiWindow), new UIPropertyMetadata(OnIsModalChanged));

        /// <summary>
        /// Identifies the <see cref="IsSelected" /> dependency property.
        /// </summary>
        public static readonly DependencyProperty IsSelectedProperty =
            DependencyProperty.Register(nameof(IsSelected), typeof(bool), typeof(MdiWindow), new UIPropertyMetadata(false));

        /// <summary>
        /// Identifies the <see cref="Title" /> dependency property.
        /// </summary>
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(nameof(Title), typeof(string), typeof(MdiWindow), new PropertyMetadata(string.Empty));

        /// <summary>
        /// Identifies the <see cref="WindowStateChanged" /> routed event.
        /// </summary>
        public static readonly RoutedEvent WindowStateChangedEvent = EventManager.RegisterRoutedEvent(
           nameof(WindowStateChanged), RoutingStrategy.Bubble, typeof(WindowStateChangedRoutedEventHandler), typeof(MdiWindow));

        /// <summary>
        /// Identifies the <see cref="WindowState" /> dependency property.
        /// </summary>
        public static readonly DependencyProperty WindowStateProperty =
            DependencyProperty.Register(nameof(WindowState), typeof(WindowState), typeof(MdiWindow), new PropertyMetadata(WindowState.Normal, OnWindowStateChanged));

        private Adorner? _adorner;

        private AdornerLayer? _ardonerLayer;

        static MdiWindow()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(MdiWindow), new FrameworkPropertyMetadata(typeof(MdiWindow)));
        }

        public void ChangeMenuButtonIcon(ImageBrush brush)
        {
            var menuButton = GetTemplateChild("PART_ButtonBar_MenuButton");
            if (menuButton is WindowButton button)
            {
                button.Icon = brush;
            }
        }

        public MdiWindow()
        {
            _ardonerLayer = AdornerLayer.GetAdornerLayer(this);
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
                    _ardonerLayer?.Remove(_adorner);
                }
                else
                {
                    if (_ardonerLayer == null)
                    {
                        _ardonerLayer = AdornerLayer.GetAdornerLayer(this);
                    }

                    if (_adorner == null)
                    {
                        _adorner = new HollowRectangleAdorner(this);
                    }
                    _ardonerLayer.Add(_adorner);
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

        public Image Tumblr { get; private set; } = new Image();

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

        internal MdiContainer? Container { get; private set; }

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
            if (GetTemplateChild("PART_ButtonBar_MenuButton") is WindowButton menuButton)
            {
                menuButton.MouseDoubleClick += CloseWindow;
            }

            if (GetTemplateChild("PART_ButtonBar_CloseButton") is WindowButton closeButton)
            {
                closeButton.Click += CloseWindow;
            }

            if (GetTemplateChild("PART_ButtonBar_MaximizeButton") is WindowButton maximizeButton)
            {
                maximizeButton.Click += ToggleMaximizeWindow;
            }

            if (GetTemplateChild("PART_ButtonBar_MinimizeButton") is WindowButton minimizeButton)
            {
                minimizeButton.Click += ToggleMinimizeWindow;
            }
            if (GetTemplateChild("PART_Tumblr") is Image image)
            {
                Tumblr = image;
            }
        }

        public void Position(bool firstAppearance = true)
        {
            if (Container != null)
            {
                if (this.DesiredSize.Height > Container.ActualHeight ||
                    this.DesiredSize.Width > Container.ActualWidth)
                {
                    this.Maximize();
                }
                else
                {
                    double left = Canvas.GetLeft(this);
                    double top = Canvas.GetTop(this);
                    if (firstAppearance)
                    {
                        left = Mouse.GetPosition(this).X - this.DesiredSize.Width / 2;
                        top = Mouse.GetPosition(this).Y - this.DesiredSize.Height / 2;
                    }
                    if (top < 0)
                    {
                        top = 0;
                    }
                    if (top + this.DesiredSize.Height > Container.ActualHeight)
                    {
                        top = Container.ActualHeight - this.DesiredSize.Height;
                    }
                    if (left < 0)
                    {
                        left = 0;
                    }
                    if (left + this.DesiredSize.Width > Container.ActualWidth)
                    {
                        left = Container.ActualWidth - this.DesiredSize.Width;
                    }
                    Canvas.SetLeft(this, left);
                    Canvas.SetTop(this, top);
                }
            }
        }

        internal void Initialize(MdiContainer container)
        {
            Container = container;
            Container.SizeChanged += OnContainerSizeChanged;
            Container.SelectionChanged += OnContainerSelectionChanged;
            LastHeight = ActualHeight;
            LastWidth = ActualWidth;
        }

        private void OnContainerSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.RemovedItems.Count > 0 && double.IsNaN(this.Width))
            {
                RemoveReferences();
                RaiseEvent(new RoutedEventArgs(ClosingEvent));
            }
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
            if (e.NewFocus is FrameworkElement element)
            {
                FrameworkElement? parent = VisualTreeExtension.FindMdiWindow(element);
                if ((e.NewFocus is MdiWindow && !Equals(e.NewFocus, this)) || (parent != null && !Equals(parent, this)))
                {
                    SetCurrentValue(IsSelectedProperty, false);
                    Panel.SetZIndex(this, 0);
                    var newWindow = (e.NewFocus is MdiWindow) ? (e.NewFocus as MdiWindow) : (parent as MdiWindow);
                    if (newWindow != null && Container != null)
                    {
                        Container.SetCurrentValue(System.Windows.Controls.Primitives.Selector.SelectedItemProperty, newWindow.DataContext);
                        newWindow.IsSelected = true;
                    }
                }
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
            if (e.NewValue == null || !((bool?)e.NewValue).HasValue) return;
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
            RemoveReferences();
            RaiseEvent(new RoutedEventArgs(ClosingEvent));
        }

        private void RemoveReferences()
        {
            if (Container != null)
            {
                Container.SizeChanged -= this.OnContainerSizeChanged;
                Container.SelectionChanged -= OnContainerSelectionChanged;
            }
            Container = null;
            _adorner = null;
            _ardonerLayer = null;
        }

        private void OnContainerSizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (WindowState == WindowState.Maximized)
            {
                Width += e.NewSize.Width - e.PreviousSize.Width;
                Height += e.NewSize.Height - e.PreviousSize.Height;
            }
            else if (WindowState == WindowState.Minimized && Container != null)
            {
                Canvas.SetTop(this, Container.ActualHeight - 24);
                LastLeft += e.NewSize.Width - e.PreviousSize.Width;
                LastLeft = Math.Max(0, LastLeft);
                LastTop += e.NewSize.Height - e.PreviousSize.Height;
                LastTop = Math.Max(0, LastTop);
                LastWidth = Math.Min(LastWidth, LastWidth + e.NewSize.Width - e.PreviousSize.Width);
                if (LastWidth < 0)
                {
                    LastWidth = Math.Min(this.DesiredSize.Width, Container.Width);
                }
                LastHeight = Math.Min(LastHeight, LastHeight + e.NewSize.Height - e.PreviousSize.Height);
                if (LastHeight < 0)
                {
                    LastHeight = Math.Min(this.DesiredSize.Height, Container.Height);
                }
            }
            else
            {
                KeepWithinContainer();
            }
        }

        internal void KeepWithinContainer()
        {
            if (Container is null)
            {
                return;
            }
            if (Canvas.GetTop(this) < 0)
            {
                Canvas.SetTop(this, 0);
            }
            if (Canvas.GetLeft(this) < 0)
            {
                Canvas.SetLeft(this, 0);
            }
            if (Canvas.GetLeft(this) + this.ActualWidth > Container.ActualWidth)
            {
                Canvas.SetLeft(this, Container.ActualWidth - (this.ActualWidth));
            }
            if (Canvas.GetTop(this) + this.ActualHeight > Container.ActualHeight)
            {
                Canvas.SetTop(this, Container.ActualHeight - (this.ActualHeight));
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