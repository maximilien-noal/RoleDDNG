using System.Windows;
using System.Windows.Controls;

namespace Hammer.MDI.Control.Panels
{
    public class ConstrainedCanvas : Canvas
    {
        protected override Size ArrangeOverride(Size arrangeSize)
        {
            System.Collections.IList list = base.InternalChildren;
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i] is UIElement internalChild)
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
                    internalChild.Arrange(new Rect(new Point(x, y), internalChild.DesiredSize));
                }
            }
            return arrangeSize;
        }

        protected override Size MeasureOverride(Size constraint)
        {
            constraint.Height = this.ActualHeight;
            constraint.Width = this.ActualWidth;
            Size availableSize = new Size(constraint.Width, constraint.Height);
            System.Collections.IList list = base.InternalChildren;
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i] is UIElement internalChild)
                {
                    internalChild.Measure(availableSize);
                }
            }
            return default;
        }
    }
}