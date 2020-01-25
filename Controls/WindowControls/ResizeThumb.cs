﻿using Hammer.MDI.Control.Extensions;
using Hammer.MdiControls.Panels;

using System;
using System.Windows;
using System.Windows.Controls.Primitives;

namespace Hammer.MDI.Control.WindowControls
{
    public sealed class ResizeThumb : Thumb
    {
        public ResizeThumb()
        {
            DragDelta += OnResizing;
        }

        private void OnResizing(object sender, DragDeltaEventArgs e)
        {
            var window = VisualTreeExtension.FindMdiWindow(this);

            if (window != null)
            {
                if (window.IsFocused == false)
                {
                    window.Focus();
                }

                window.Height = window.ActualHeight;
                window.Width = window.ActualWidth;

                switch (VerticalAlignment)
                {
                    case VerticalAlignment.Bottom:
                        var deltaVertical = Math.Min(-e.VerticalChange, window.ActualHeight - window.MinHeight);
                        window.Height -= deltaVertical;
                        break;

                    case VerticalAlignment.Top:
                        deltaVertical = Math.Min(e.VerticalChange, window.ActualHeight - window.MinHeight);
                        AutoResizeCanvas.SetTop(window, AutoResizeCanvas.GetTop(window) + deltaVertical);
                        window.Height -= deltaVertical;
                        break;

                    default:
                        throw new Exception("Unexpected Case");
                }

                switch (HorizontalAlignment)
                {
                    case HorizontalAlignment.Left:
                        var deltaHorizontal = Math.Min(e.HorizontalChange, window.ActualWidth - window.MinWidth);
                        AutoResizeCanvas.SetLeft(window, AutoResizeCanvas.GetLeft(window) + deltaHorizontal);
                        window.Width -= deltaHorizontal;
                        break;

                    case HorizontalAlignment.Right:
                        deltaHorizontal = Math.Min(-e.HorizontalChange, window.ActualWidth - window.MinWidth);
                        window.Width -= deltaHorizontal;
                        break;

                    case HorizontalAlignment.Center:
                        break;

                    case HorizontalAlignment.Stretch:
                        break;

                    default:
                        break;
                }
            }

            e.Handled = true;
        }
    }
}