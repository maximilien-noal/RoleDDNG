using System.Windows;
using System.Windows.Controls;

namespace Hammer.MdiControls.Panels
{
    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    public class AutoResizeCanvas : Canvas
    {
        private Size _previousArrangeSize = new Size();

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        protected override Size MeasureOverride(Size constraint)
        {
            var maxRightBottomChildSize = new Size();
            foreach (UIElement internalChild in InternalChildren)
            {
                if (internalChild == null) { continue; }
                internalChild.Measure(RenderSize);
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
            if (InternalChildren.Count == 0)
            {
                return new Size();
            }
            if (RenderSize.Width > maxRightBottomChildSize.Width && RenderSize.Height > maxRightBottomChildSize.Height)
            {
                return RenderSize;
            }
            return maxRightBottomChildSize;
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

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        protected override Size ArrangeOverride(Size arrangeSize)
        {
            var maxChildRightBottom = GetMaxRightBottomChild();
            if (arrangeSize.Width < maxChildRightBottom.Width || arrangeSize.Height < maxChildRightBottom.Height)
            {
                _previousArrangeSize = maxChildRightBottom;
                InvalidateMeasure();
            }
            else if (_previousArrangeSize.Width > maxChildRightBottom.Width || _previousArrangeSize.Height > maxChildRightBottom.Height)
            {
                _previousArrangeSize = maxChildRightBottom;
                InvalidateMeasure();
            }

            var calculatedBase = base.ArrangeOverride(arrangeSize);

            if (calculatedBase.Height < 0 || calculatedBase.Width < 0)
            {
                calculatedBase.Height = 0;
                calculatedBase.Width = 0;
                InvalidateArrange();
            }
            return calculatedBase;
        }
    }
}