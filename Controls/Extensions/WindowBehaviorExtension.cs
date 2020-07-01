using System.Windows;
using System.Windows.Controls;

using Hammer.MDI.Control.Panels;

namespace Hammer.MDI.Control.Extensions
{
    internal static class WindowBehaviorExtension
    {
        public static void Maximize(this MdiWindow window)
        {
            SaveLastSize(window);
            ConstrainedCanvas.SetTop(window, 0.0);
            ConstrainedCanvas.SetLeft(window, 0.0);
            if (window.Container != null)
            {
                window.SetCurrentValue(FrameworkElement.HeightProperty, window.Container.ActualHeight);
                window.SetCurrentValue(FrameworkElement.WidthProperty, window.Container.ActualWidth);
            }

            window.SetCurrentValue(MdiWindow.WindowStateProperty, WindowState.Maximized);
        }

        public static void Minimize(this MdiWindow window)
        {
            var index = 0;
            if (window.Container != null)
            {
                index = window.Container.MinimizedWindowsCount;
            }
            SaveLastSize(window);
            window.Tumblr.SetCurrentValue(Image.SourceProperty, window.CreateSnapshot());
            window.SetCurrentValue(FrameworkElement.WidthProperty, (double)128);
            if (window.Container != null)
            {
                ConstrainedCanvas.SetTop(window, window.Container.ActualHeight - 24);
            }
            ConstrainedCanvas.SetLeft(window, index * 205);

            window.SetCurrentValue(MdiWindow.WindowStateProperty, WindowState.Minimized);
        }

        public static void Normalize(this MdiWindow window)
        {
            window.SetCurrentValue(FrameworkElement.HeightProperty, window.LastHeight);
            window.SetCurrentValue(FrameworkElement.WidthProperty, window.LastWidth);
            ConstrainedCanvas.SetTop(window, window.LastTop);
            ConstrainedCanvas.SetLeft(window, window.LastLeft);

            window.SetCurrentValue(MdiWindow.WindowStateProperty, WindowState.Normal);

            //window.KeepWithinContainer();
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
            window.LastLeft = ConstrainedCanvas.GetLeft(window);
            window.LastTop = ConstrainedCanvas.GetTop(window);
            window.LastWidth = window.ActualWidth;
            window.LastHeight = window.ActualHeight;
        }
    }
}