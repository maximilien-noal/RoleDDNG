using System.Windows;

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

            window.Height = window.Container.ActualHeight;
            window.Width = window.Container.ActualWidth;

            window.WindowState = WindowState.Maximized;
        }

        public static void Minimize(this MdiWindow window)
        {
            var index = window.Container.MinimizedWindowsCount;

            SaveLastSize(window);
            window.Tumblr.Source = window.CreateSnapshot();
            window.Width = 128;
            AutoResizeCanvas.SetTop(window, window.Container.ActualHeight - 32);
            AutoResizeCanvas.SetLeft(window, index * 205);

            window.WindowState = WindowState.Minimized;
        }

        public static void Normalize(this MdiWindow window)
        {
            window.Height = window.LastHeight;
            window.Width = window.LastWidth;
            AutoResizeCanvas.SetTop(window, window.LastTop);
            AutoResizeCanvas.SetLeft(window, window.LastLeft);

            window.WindowState = WindowState.Normal;
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
            window.LastLeft = AutoResizeCanvas.GetLeft(window);
            window.LastTop = AutoResizeCanvas.GetTop(window);
            window.LastWidth = window.ActualWidth;
            window.LastHeight = window.ActualHeight;
        }
    }
}