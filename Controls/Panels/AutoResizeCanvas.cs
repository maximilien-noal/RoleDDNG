using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

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
            Size availableSize = new Size(ActualWidth, ActualHeight);
            var parent = FindParent<ScrollViewer>(this);
            if (parent != null)
            {
                availableSize = new Size(parent.ActualWidth, parent.ActualHeight);
            }
            foreach (UIElement internalChild in base.InternalChildren)
            {
                internalChild?.Measure(availableSize);
            }
            return availableSize;
        }

        private static T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            //get parent item
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);

            //we've reached the end of the tree
            if (parentObject == null) return null;

            //check if the parent matches the type we're looking for
            if (parentObject is T parent)
                return parent;
            else
                return FindParent<T>(parentObject);
        }
    }
}