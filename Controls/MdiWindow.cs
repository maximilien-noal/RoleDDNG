﻿using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;

using Hammer.MDI.Control.Events;
using Hammer.MDI.Control.Extensions;
using Hammer.MDI.Control.WindowControls;
using Hammer.MdiControls.Panels;

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
    public sealed class MdiWindow : ContentControl
    {
        public static readonly RoutedEvent ClosingEvent = EventManager.RegisterRoutedEvent(
            nameof(Closing), RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(MdiWindow));

        public static readonly RoutedEvent FocusChangedEvent = EventManager.RegisterRoutedEvent(
           nameof(FocusChanged), RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(MdiWindow));

        public static readonly DependencyProperty HasDropShadowProperty =
            DependencyProperty.Register(nameof(HasDropShadow), typeof(bool), typeof(MdiWindow), new UIPropertyMetadata(true));

        public static readonly DependencyProperty IsModalProperty =
            DependencyProperty.Register(nameof(IsModal), typeof(bool?), typeof(MdiWindow), new UIPropertyMetadata(IsModalChangedCallback));

        public static readonly DependencyProperty IsSelectedProperty =
            DependencyProperty.Register(nameof(IsSelected), typeof(bool), typeof(MdiWindow), new UIPropertyMetadata(false));

        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(nameof(Title), typeof(string), typeof(MdiWindow), new PropertyMetadata(string.Empty));

        public static readonly RoutedEvent WindowStateChangedEvent = EventManager.RegisterRoutedEvent(
           nameof(WindowStateChanged), RoutingStrategy.Bubble, typeof(WindowStateChangedRoutedEventHandler), typeof(MdiWindow));

        public static readonly DependencyProperty WindowStateProperty =
            DependencyProperty.Register(nameof(WindowState), typeof(WindowState), typeof(MdiWindow), new PropertyMetadata(WindowState.Normal, OnWindowStateChanged));

        private WindowButton _closeButton;

        private WindowState _lastState;

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

        public bool IsModal
        {
            get
            {
                return (bool)GetValue(IsModalProperty);
            }

            set
            {
                if (value)
                {
                    if (_myAdornerLayer == null)
                        _myAdornerLayer = AdornerLayer.GetAdornerLayer(this);
                    if (_myAdorner == null)
                    {
                        _myAdorner = new HollowRectangleAdorner(this);
                    }
                    _myAdornerLayer.Add(_myAdorner);
                }
                else
                {
                    _myAdornerLayer?.Remove(_myAdorner);
                }
                if (Container != null)
                    Container.IsModal = value;

                SetValue(IsModalProperty, value);
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
                _lastState = (WindowState)GetValue(WindowStateProperty);
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

        public bool IsLastStateMaximized() => _lastState == WindowState.Maximized;

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

            SetValue(AutoResizeCanvas.LeftProperty, left);
            SetValue(AutoResizeCanvas.TopProperty, top);
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

            IsSelected = true;
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
                IsSelected = false;
                Panel.SetZIndex(this, 0);
                var newWindow = (e.NewFocus is MdiWindow) ? (e.NewFocus as MdiWindow) : (parent as MdiWindow);
                Container.SetValue(MdiContainer.SelectedItemProperty, newWindow.DataContext);
                newWindow.IsSelected = true;
            }
        }

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            IsSelected = true;
            Panel.SetZIndex(this, 2);
            RaiseEvent(new RoutedEventArgs(FocusChangedEvent, DataContext));

            Focus();
        }

        private static void IsModalChangedCallback(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (e.NewValue == null) return;
            ((MdiWindow)d).IsModal = (bool)e.NewValue;
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
                AutoResizeCanvas.SetTop(this, Container.ActualHeight - 32);
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