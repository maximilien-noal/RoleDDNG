using System.Windows;
using System.Windows.Controls;

namespace Hammer.MDI.Control.Extensions
{
    internal static class WindowBehaviorExtension
    {
        public static void Maximize(this MdiWindow window)
        {
            SaveLastSize(window);
            Canvas.SetTop(window, 0.0);
            Canvas.SetLeft(window, 0.0);

            window.SetCurrentValue(FrameworkElement.HeightProperty, window.Container.ActualHeight);
            window.SetCurrentValue(FrameworkElement.WidthProperty, window.Container.ActualWidth);

            window.SetCurrentValue(MdiWindow.WindowStateProperty, WindowState.Maximized);
        }

        public static void Minimize(this MdiWindow window)
        {
            var index = window.Container.MinimizedWindowsCount;

            SaveLastSize(window);
            window.Tumblr.SetCurrentValue(System.Windows.Controls.Image.SourceProperty, window.CreateSnapshot());
            window.SetCurrentValue(FrameworkElement.WidthProperty, (double)128);
            Canvas.SetTop(window, window.Container.ActualHeight - 32);
            Canvas.SetLeft(window, index * 205);

            window.SetCurrentValue(MdiWindow.WindowStateProperty, WindowState.Minimized);
        }

        public static void Normalize(this MdiWindow window)
        {
            window.SetCurrentValue(FrameworkElement.HeightProperty, window.LastHeight);
            window.SetCurrentValue(FrameworkElement.WidthProperty, window.LastWidth);
            Canvas.SetTop(window, window.LastTop);
            Canvas.SetLeft(window, window.LastLeft);

            window.SetCurrentValue(MdiWindow.WindowStateProperty, WindowState.Normal);
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

        private static void SaveLastSize(MdiWindow window)
        {
            if (window.WindowState != WindowState.Normal)
            {
                return;
            }
            window.LastLeft = Canvas.GetLeft(window);
            window.LastTop = Canvas.GetTop(window);
            window.LastWidth = window.ActualWidth;
            window.LastHeight = window.ActualHeight;
        }
    }
}