using System.Windows;
using System.Windows.Controls;

using Hammer.MDI.Control;

namespace Hammer.MdiControls.Panels
{
    /// <summary>
    /// A <see cref="Canvas"/> where its size is dependant of its children
    /// </summary>
    public class AutoResizeCanvas : Canvas
    {
        private Size _lastMeasure = new Size();

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        protected override Size MeasureOverride(Size constraint)
        {
            var maxRightBottomChildSize = new Size();
            var foundMaximizedWindow = false;
            foreach (UIElement internalChild in InternalChildren)
            {
                if (internalChild == null) { continue; }
                if (internalChild is MdiWindow window && window.WindowState == WindowState.Maximized)
                {
                    foundMaximizedWindow = true;
                }

                internalChild.Measure(RenderSize);
                if (foundMaximizedWindow == false)
                {
                    var childRight = internalChild.DesiredSize.Width + GetLeft(internalChild);
                    var childBottom = internalChild.DesiredSize.Height + GetTop(internalChild);
                    if (childRight > maxRightBottomChildSize.Width)
                    {
                        maxRightBottomChildSize.Width = childRight;
                    }
                    if (childBottom > maxRightBottomChildSize.Height)
                    {
                        maxRightBottomChildSize.Height = childBottom;
                    }
                }
            }
            if (foundMaximizedWindow)
            {
                _lastMeasure = RenderSize;
                return new Size();
            }
            _lastMeasure = maxRightBottomChildSize;
            return _lastMeasure;
        }

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        protected override Size ArrangeOverride(Size arrangeSize)
        {
            if (IsLastMeasureObsolete(arrangeSize))
            {
                InvalidateMeasure();
            }
            if (IsAnyInternalChildMaximized())
            {
                foreach (UIElement children in InternalChildren)
                {
                    if (children == null)
                    {
                        continue;
                    }
                    if (children is MdiWindow window && window.WindowState == WindowState.Normal)
                    {
                        window.WindowState = WindowState.Maximized;
                    }
                }
            }
            return base.ArrangeOverride(arrangeSize);
        }

        private bool IsLastMeasureObsolete(Size arrangeSize)
        {
            if (InternalChildren.Count == 0)
            {
                return false;
            }
            if (IsAnyInternalChildMaximized())
            {
                return IsLastMeasureEqualToArrangeSize(arrangeSize) == false;
            }
            var maxChildRightBottom = GetMaxRightBottomChild();
            return _lastMeasure.Width != maxChildRightBottom.Width || _lastMeasure.Height != maxChildRightBottom.Height;
        }

        private bool IsLastMeasureEqualToArrangeSize(Size arrangeSize)
        {
            return _lastMeasure.Width == arrangeSize.Width && _lastMeasure.Height == arrangeSize.Height;
        }

        private bool IsAnyInternalChildMaximized()
        {
            var anyMaximized = false;
            foreach (UIElement internalChild in InternalChildren)
            {
                if (internalChild == null) { continue; }
                if (internalChild is MdiWindow window && window.WindowState == WindowState.Maximized)
                {
                    return true;
                }
            }
            return anyMaximized;
        }

        private Size GetMaxRightBottomChild()
        {
            var maxRightBottomChildSize = new Size();

            foreach (UIElement internalChild in InternalChildren)
            {
                if (internalChild == null) { continue; }
                var childRight = internalChild.DesiredSize.Width + GetLeft(internalChild);
                var childBottom = internalChild.DesiredSize.Height + GetTop(internalChild);
                if (childRight > maxRightBottomChildSize.Width)
                {
                    maxRightBottomChildSize.Width = childRight;
                }
                if (childBottom > maxRightBottomChildSize.Height)
                {
                    maxRightBottomChildSize.Height = childBottom;
                }
            }
            return maxRightBottomChildSize;
        }
    }
}