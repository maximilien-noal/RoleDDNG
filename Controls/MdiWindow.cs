using Hammer.MDI.Control.Events;
using Hammer.MDI.Control.Extensions;
using Hammer.MDI.Control.WindowControls;

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

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
        public RenderTargetBitmap CreateSnapshot()
        {
            var bitmap = new RenderTargetBitmap((int)Math.Round(ActualWidth), (int)Math.Round(ActualHeight), 96, 96, PixelFormats.Default);
            var drawingVisual = new DrawingVisual();
            using (var context = drawingVisual.RenderOpen())
            {
                var brush = new VisualBrush(this);
                context.DrawRectangle(brush, null, new Rect(new Point(), new Size(ActualWidth, ActualHeight)));
                context.Close();
            }

            bitmap.Render(drawingVisual);

            return bitmap;
        }

        internal void Maximize()
        {
            SaveLastSizeAndPosition();
            Canvas.SetTop(this, 0.0);
            Canvas.SetLeft(this, 0.0);
            if (Container != null)
            {
                SetCurrentValue(HeightProperty, Container.ActualHeight);
                SetCurrentValue(WidthProperty, Container.ActualWidth);
            }

            SetCurrentValue(WindowStateProperty, WindowState.Maximized);
        }

        internal void Minimize()
        {
            if (Container is null)
            {
                return;
            }
            var index = Container.MinimizedWindowsCount;
            SetCurrentValue(MinWidthProperty, 205d);
            SetCurrentValue(WidthProperty, 205d);
            SaveLastSizeAndPosition();
            Tumblr.SetCurrentValue(Image.SourceProperty, CreateSnapshot());
            Canvas.SetTop(this, Container.ActualHeight - ChromeHeight);
            Canvas.SetLeft(this, index * 205d);
            SetCurrentValue(WindowStateProperty, WindowState.Minimized);
        }

        private double GetYAxisValueWithinContainer(double calculatedTop)
        {
            if (Container is null || calculatedTop < 0)
            {
                return 0;
            }
            else if (calculatedTop + ActualHeight > Container.ActualHeight)
            {
                return Container.ActualHeight - ActualHeight;
            }
            return calculatedTop;
        }

        private double GetXAxisValueWithinContainer(double calculatedLeft)
        {
            if (Container is null || calculatedLeft < 0)
            {
                return 0;
            }
            else if (calculatedLeft + ActualWidth > Container.ActualWidth)
            {
                return Container.ActualWidth - ActualWidth;
            }

            return calculatedLeft;
        }

        internal void GetChromeUnderMouse()
        {
            if (Container is null)
            {
                return;
            }
            Canvas.SetTop(this, GetYAxisValueWithinContainer(Mouse.GetPosition(Container).Y - ChromeHeight / 2));
            Canvas.SetLeft(this, GetXAxisValueWithinContainer(Mouse.GetPosition(Container).X - ChromeWidth * 2));
            SaveLastSizeAndPosition();
        }

        internal void Restore()
        {
            SetCurrentValue(HeightProperty, LastHeight);
            SetCurrentValue(MinWidthProperty, LastWidth);
            SetCurrentValue(WidthProperty, LastWidth);
            Canvas.SetTop(this, LastTop);
            Canvas.SetLeft(this, LastLeft);
            SetCurrentValue(WindowStateProperty, WindowState.Normal);
            SaveLastSizeAndPosition();
        }

        internal void ToggleMaximize()
        {
            if (WindowState == WindowState.Maximized)
            {
                Restore();
            }
            else
            {
                Maximize();
            }
        }

        public void ToggleMinimize()
        {
            if (WindowState != WindowState.Minimized)
            {
                Minimize();
            }
            else
            {
                switch (PreviousWindowState)
                {
                    case WindowState.Maximized:
                        Maximize();
                        break;

                    case WindowState.Normal:
                        Restore();
                        break;

                    default:
                        break;
                }
            }
        }

        internal void SaveFirstSize()
        {
            if (WindowState != WindowState.Normal || LastWidth != 0 || LastHeight != 0)
            {
                return;
            }
            LastLeft = Canvas.GetLeft(this);
            LastTop = Canvas.GetTop(this);
            LastWidth = ActualWidth;
            LastHeight = ActualHeight;
        }

        private bool _unlockPass;

        internal void UnlockDimensionsOnce()
        {
            if (_unlockPass)
            {
                return;
            }
            _unlockPass = true;
            SetCurrentValue(MinWidthProperty, ActualWidth);
            SetCurrentValue(MaxWidthProperty, double.PositiveInfinity);
            SetCurrentValue(MinHeightProperty, ActualHeight);
            SetCurrentValue(MaxHeightProperty, double.PositiveInfinity);
        }

        private bool _lockPass;

        internal void LockDimensionsOnce()
        {
            if (_lockPass)
            {
                return;
            }
            _lockPass = true;
            SetCurrentValue(MinWidthProperty, ActualWidth);
            SetCurrentValue(MaxWidthProperty, ActualWidth);
            SetCurrentValue(MinHeightProperty, ActualHeight);
            SetCurrentValue(MaxHeightProperty, ActualHeight);
        }

        internal void SaveLastSizeAndPosition()
        {
            if (WindowState != WindowState.Normal)
            {
                return;
            }
            LastLeft = Canvas.GetLeft(this);
            LastTop = Canvas.GetTop(this);
            LastWidth = ActualWidth;
            LastHeight = ActualHeight;
        }

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

        private AdornerLayer? _myAdornerLayer;

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

        private FrameworkElement? _windowChrome;

        public double ChromeHeight => _windowChrome is null ? 24 : _windowChrome.ActualHeight;

        public double ChromeWidth => _windowChrome is null ? 24 : _windowChrome.ActualHeight;

        public override void OnApplyTemplate()
        {
            if (GetTemplateChild("PART_ButtonBar") is FrameworkElement windowChrome)
            {
                _windowChrome = windowChrome;
            }

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

        public void Position()
        {
            if (Container == null)
            {
                return;
            }
            ResizeToAvailableSpace();
            double left = Container.ActualWidth / 4 - (DesiredSize.Width / 2);
            double top = Container.ActualHeight / 4 - (DesiredSize.Height / 2);
            Canvas.SetLeft(this, GetXAxisValueWithinContainer(left));
            Canvas.SetTop(this, GetYAxisValueWithinContainer(top));
        }

        public void ResizeToAvailableSpace()
        {
            if (Container == null)
            {
                return;
            }

            if
                (
                Math.Max(0, Canvas.GetTop(this)) + DesiredSize.Height > Container.ActualHeight ||
                Math.Max(0, Canvas.GetLeft(this)) + DesiredSize.Width > Container.ActualWidth)
            {
                Rect availableRect = new Rect(
                    Canvas.GetLeft(this), Canvas.GetTop(this),
                    Math.Min(DesiredSize.Width, Container.ActualWidth - Math.Max(0, Canvas.GetLeft(this))),
                    Math.Min(DesiredSize.Height + ChromeHeight, Container.ActualHeight - Math.Max(0, Canvas.GetTop(this))));
                SetCurrentValue(WidthProperty, availableRect.Width);
                SetCurrentValue(HeightProperty, availableRect.Height);
            }
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
                        Container.SetCurrentValue(Selector.SelectedItemProperty, newWindow.DataContext);
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
            if (Container != null)
            {
                Container.SizeChanged -= this.OnContainerSizeChanged;
            }
            RaiseEvent(new RoutedEventArgs(ClosingEvent));
        }

        private void OnContainerSizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (WindowState == WindowState.Maximized)
            {
                SetCurrentValue(WidthProperty, e.NewSize.Width);
                SetCurrentValue(HeightProperty, e.NewSize.Height);
            }
            else if (WindowState == WindowState.Minimized && Container != null)
            {
                Canvas.SetTop(this, Container.ActualHeight - ChromeHeight);
                LastLeft += e.NewSize.Width - e.PreviousSize.Width;
                LastLeft = Math.Max(0, LastLeft);
                LastTop += e.NewSize.Height - e.PreviousSize.Height;
                LastTop = Math.Max(0, LastTop);
                if (LastLeft + LastWidth > e.NewSize.Width)
                {
                    LastWidth = Math.Min(LastWidth, LastWidth + e.NewSize.Width - e.PreviousSize.Width);
                }
                if (LastWidth < 0)
                {
                    LastWidth = Math.Min(this.DesiredSize.Width, Container.Width);
                }
                if (LastTop + LastHeight > e.NewSize.Height)
                {
                    LastHeight = Math.Min(LastHeight, LastHeight + e.NewSize.Height - e.PreviousSize.Height);
                }
            }
            else
            {
                KeepWithinContainer();
                ResizeToAvailableSpace();
            }
        }

        private void KeepWithinContainer()
        {
            if (Container is null)
            {
                return;
            }
            Canvas.SetLeft(this, GetXAxisValueWithinContainer(Canvas.GetLeft(this)));
            Canvas.SetTop(this, GetYAxisValueWithinContainer(Canvas.GetTop(this)));
        }

        internal void FocusOnFirstFocusableChild()
        {
            var child = FindChild(this);
            if (child is FrameworkElement fe)
            {
                fe.Focus();
            }
        }

        private static FrameworkElement? FindChild(DependencyObject parent)
        {
            if (parent == null)
            {
                return null;
            }

            int childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is FrameworkElement fe && fe.Focusable)
                {
                    return fe;
                }
                else
                {
                    var foundChild = FindChild(child);
                    if (foundChild != null)
                    {
                        return foundChild;
                    }
                }
            }

            return null;
        }

        private void ToggleMaximizeWindow(object sender, RoutedEventArgs e)
        {
            Focus();
            ToggleMaximize();
        }

        private void ToggleMinimizeWindow(object sender, RoutedEventArgs e)
        {
            Focus();
            ToggleMinimize();
        }
    }
}