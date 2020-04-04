using System.Windows;
using System.Windows.Controls;

namespace Hammer.MdiControls.Panels
{
    /// <summary>
    /// A <see cref="Canvas" /> where its size is dependant of its parent
    /// </summary>
    public class AutoResizeCanvas : Canvas
    {
        /// <summary>
        /// <inheritdoc />
        /// </summary>
        protected override Size ArrangeOverride(Size arrangeSize)
        {
            return base.ArrangeOverride(arrangeSize);
        }

        /// <summary>
        /// <inheritdoc />
        /// </summary>
        protected override Size MeasureOverride(Size constraint)
        {
            Size availableSize = new Size(SystemParameters.VirtualScreenHeight, SystemParameters.VirtualScreenWidth);
            foreach (UIElement internalChild in base.InternalChildren)
            {
                internalChild?.Measure(availableSize);
            }
            return availableSize;
        }
    }
}