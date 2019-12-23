using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Hammer.MDI.Control;

namespace Hammer.MdiControls.Panels
{
    /// <summary>
    /// <inheritdoc/>
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
                return _lastMeasure;
            }
            _lastMeasure = maxRightBottomChildSize;
            return _lastMeasure;
        }

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        protected override Size ArrangeOverride(Size arrangeSize)
        {
            if (arrangeSize.Height < 0 || arrangeSize.Width < 0)
            {
                arrangeSize.Height = 0;
                arrangeSize.Width = 0;
                InvalidateArrange();
            }
            if (IsLastMeasureObsolete(arrangeSize))
            {
                InvalidateMeasure();
            }
            foreach (UIElement internalChild in InternalChildren)
            {
                if (internalChild != null)
                {
                    double x = 0.0;
                    double y = 0.0;
                    double left = GetLeft(internalChild);
                    if (!double.IsNaN(left))
                    {
                        x = left;
                    }
                    else
                    {
                        double right = GetRight(internalChild);
                        if (!double.IsNaN(right))
                        {
                            x = arrangeSize.Width - internalChild.DesiredSize.Width - right;
                        }
                    }
                    double top = GetTop(internalChild);
                    if (!double.IsNaN(top))
                    {
                        y = top;
                    }
                    else
                    {
                        double bottom = GetBottom(internalChild);
                        if (!double.IsNaN(bottom))
                        {
                            y = arrangeSize.Height - internalChild.DesiredSize.Height - bottom;
                        }
                    }
                    if (internalChild is MdiWindow window && window.WindowState == WindowState.Maximized)
                    {
                        internalChild.Arrange(new Rect(new Point(0, 0), RenderSize));
                    }
                    else
                    {
                        internalChild.Arrange(new Rect(new Point(x, y), internalChild.DesiredSize));
                    }
                }
            }
            return arrangeSize;
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