using Hammer.MDI.Control.Extensions;

using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace Hammer.MDI.Control.WindowControls
{
    internal sealed class ResizeThumb : Thumb
    {
        public ResizeThumb()
        {
            DragDelta += OnResizing;
        }

        private void OnResizing(object sender, DragDeltaEventArgs e)
        {
            var window = VisualTreeExtension.FindMdiWindow(this);

            if (window == null)
            {
                return;
            }
            else
            {
                if (!window.IsFocused)
                {
                    window.Focus();
                }

                window.Height = window.ActualHeight;
                window.Width = window.ActualWidth;

                switch (VerticalAlignment)
                {
                    case VerticalAlignment.Bottom:
                        var deltaVertical = Math.Min(-e.VerticalChange, window.ActualHeight - window.MinHeight);
                        window.Height = Math.Abs(window.Height - deltaVertical);
                        break;

                    case VerticalAlignment.Top:
                        deltaVertical = Math.Min(e.VerticalChange, window.ActualHeight - window.MinHeight);
                        Canvas.SetTop(window, Canvas.GetTop(window) + deltaVertical);
                        window.Height = Math.Abs(window.Height - deltaVertical);
                        break;

                    default:
                        break;
                }

                switch (HorizontalAlignment)
                {
                    case HorizontalAlignment.Left:
                        var deltaHorizontal = Math.Min(e.HorizontalChange, window.ActualWidth - window.MinWidth);
                        Canvas.SetLeft(window, Canvas.GetLeft(window) + deltaHorizontal);
                        window.Width = Math.Abs(window.Width - deltaHorizontal);
                        break;

                    case HorizontalAlignment.Right:
                        deltaHorizontal = Math.Min(-e.HorizontalChange, window.ActualWidth - window.MinWidth);
                        window.Width = Math.Abs(window.Width - deltaHorizontal);
                        break;

                    case HorizontalAlignment.Center:
                        break;

                    case HorizontalAlignment.Stretch:
                        break;

                    default:
                        break;
                }
            }
            var topCrop = Canvas.GetTop(window);
            if (topCrop < 0)
            {
                Canvas.SetTop(window, 0d);
                window.Height += topCrop;
            }
            e.Handled = true;
        }
    }
}