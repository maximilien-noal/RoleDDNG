using System.Windows;
using System.Windows.Controls;

namespace Hammer.MdiControls.Panels
{
    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    public class AutoResizeCanvas : Canvas
    {
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
            return maxRightBottomChildSize;
        }

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        protected override Size ArrangeOverride(Size arrangeSize)
        {
            var maxChildRightBottom = GetMaxRightBottomChild();
            if (RenderSize.Width != maxChildRightBottom.Width || RenderSize.Height != maxChildRightBottom.Height)
            {
                if (maxChildRightBottom.Height != 0 && maxChildRightBottom.Width != 0)
                {
                    InvalidateMeasure();
                }
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