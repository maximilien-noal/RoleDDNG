using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Animation;

using Hammer.MdiControls.Panels;

namespace Hammer.MDI.Control.Extensions
{
    internal static class WindowBehaviorExtension
    {
        public static void Maximize(this MdiWindow window)
        {
            SaveLastSize(window);
            AutoResizeCanvas.SetTop(window, 0.0);
            AutoResizeCanvas.SetLeft(window, 0.0);

            AnimateResize(window, window.Container.ActualWidth, window.Container.ActualHeight, true);

            window.WindowState = WindowState.Maximized;
        }

        public static void Minimize(this MdiWindow window)
        {
            var index = window.Container.MinimizedWindowsCount;

            SaveLastSize(window);
            AutoResizeCanvas.SetTop(window, window.Container.ActualHeight - 32);
            AutoResizeCanvas.SetLeft(window, index * 205);

            RemoveWindowLock(window);
            AnimateResize(window, 200, 32, true);

            window.WindowState = WindowState.Minimized;

            window.Tumblr.Source = window.CreateSnapshot();
        }

        public static void Normalize(this MdiWindow window)
        {
            AutoResizeCanvas.SetTop(window, window.LastTop);
            AutoResizeCanvas.SetLeft(window, window.LastLeft);

            AnimateResize(window, window.LastWidth, window.LastHeight, false);

            window.WindowState = WindowState.Normal;
        }

        public static void RemoveWindowLock(this MdiWindow window)
        {
            window.BeginAnimation(FrameworkElement.WidthProperty, null);
            window.BeginAnimation(FrameworkElement.HeightProperty, null);
        }

        public static void ToggleMaximize(this MdiWindow window)
        {
            if (window.WindowState == WindowState.Maximized)
            {
                window.Normalize();
            }
            else
            {
                window.Maximize();
            }
        }

        public static void ToggleMinimize(this MdiWindow window)
        {
            if (window is null)
            {
                return;
            }

            if (window.WindowState != WindowState.Minimized)
            {
                window.Minimize();
            }
            else
            {
                switch (window.PreviousWindowState)
                {
                    case WindowState.Maximized:
                        window.Maximize();
                        break;

                    case WindowState.Normal:
                        window.Normalize();
                        break;

                    default:
                        break;
                }
            }
        }

        private static void AnimateResize(MdiWindow window, double newWidth, double newHeight, bool lockWindow)
        {
            window.LayoutTransform = new ScaleTransform();

            var widthAnimation = new DoubleAnimation(window.ActualWidth, newWidth, new Duration(TimeSpan.FromMilliseconds(10)));
            var heightAnimation = new DoubleAnimation(window.ActualHeight, newHeight, new Duration(TimeSpan.FromMilliseconds(10)));

            if (lockWindow == false)
            {
                widthAnimation.Completed += (s, e) => window.BeginAnimation(FrameworkElement.WidthProperty, null);
                heightAnimation.Completed += (s, e) => window.BeginAnimation(FrameworkElement.HeightProperty, null);
            }

            window.BeginAnimation(FrameworkElement.WidthProperty, widthAnimation, HandoffBehavior.Compose);
            window.BeginAnimation(FrameworkElement.HeightProperty, heightAnimation, HandoffBehavior.Compose);
        }

        private static void SaveLastSize(MdiWindow window)
        {
            if (window.WindowState == WindowState.Normal)
            {
                window.LastLeft = AutoResizeCanvas.GetLeft(window);
                window.LastTop = AutoResizeCanvas.GetTop(window);
                window.LastWidth = window.ActualWidth;
                window.LastHeight = window.ActualHeight;
            }
        }
    }
}