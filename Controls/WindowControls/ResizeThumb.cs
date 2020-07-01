﻿using Hammer.MDI.Control.Extensions;
using Hammer.MDI.Control.Panels;

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
            var topCrop = ConstrainedCanvas.GetTop(window);
            if (topCrop < 0)
            {
                ConstrainedCanvas.SetTop(window, 0d);
                window.Height += topCrop;
            }

            //var leftCrop = ConstrainedCanvas.GetLeft(window);
            //if (leftCrop < 0)
            //{
            //    ConstrainedCanvas.SetLeft(window, 0d);
            //    window.Width += leftCrop;
            //}
            //var rightCrop = ConstrainedCanvas.GetRight(window);
            //if (double.IsNaN(rightCrop))
            //{
            //    var containerWidth = window.Container?.ActualWidth;
            //    if (containerWidth != null)
            //    {
            //        ConstrainedCanvas.SetRight(window, containerWidth.Value - ConstrainedCanvas.GetLeft(window));
            //    }
            //}
            //var bottomCrop = ConstrainedCanvas.GetBottom(window);

            //if (double.IsNaN(bottomCrop))
            //{
            //    var containerWidth = window.Container?.ActualWidth;
            //    if (containerWidth != null)
            //    {
            //        ConstrainedCanvas.SetBottom(window, containerWidth.Value - ConstrainedCanvas.GetTop(window));
            //    }
            //}
            e.Handled = true;
        }
    }
}