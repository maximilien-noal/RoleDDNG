﻿using Hammer.MDI.Control.Extensions;

using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace Hammer.MDI.Control.WindowControls
{
    internal sealed class MoveThumb : Thumb
    {
        static MoveThumb()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(MoveThumb), new FrameworkPropertyMetadata(typeof(MoveThumb)));
        }

        public MoveThumb()
        {
            DragDelta += OnMoveThumbDragDelta;
        }

        protected override void OnMouseDoubleClick(MouseButtonEventArgs e)
        {
            var window = VisualTreeExtension.FindMdiWindow(this);

            if (window != null && window.Container != null)
            {
                switch (window.WindowState)
                {
                    case WindowState.Maximized:
                        window.Normalize();
                        break;

                    case WindowState.Normal:
                        window.Maximize();
                        break;

                    case WindowState.Minimized:
                        window.Normalize();
                        break;

                    default:
                        break;
                }
            }

            e.Handled = true;
        }

        protected override void OnMouseUp(MouseButtonEventArgs e)
        {
            if (e is null)
            {
                return;
            }

            var window = VisualTreeExtension.FindMdiWindow(this);
            window?.KeepWithinContainer();
            base.OnMouseUp(e);
        }

        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            if (e is null)
            {
                return;
            }

            var window = VisualTreeExtension.FindMdiWindow(this);

            if (window != null)
            {
                window.DoFocus(e);
            }

            base.OnMouseDown(e);
        }

        private void OnMoveThumbDragDelta(object sender, DragDeltaEventArgs e)
        {
            var window = VisualTreeExtension.FindMdiWindow(this);

            if (window == null)
            {
                return;
            }
            if (window.WindowState == WindowState.Maximized)
            {
                window.Normalize();
            }

            if (window.WindowState == WindowState.Normal)
            {
                window.SetChromeCenterOnMouse();
            }
        }
    }
}