using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace Hammer.MDI.Control.Panels
{
    public class ScrollableCanvas : Canvas
    {
        protected override Size MeasureOverride(Size constraint)
        {
            var defaultValue = base.MeasureOverride(constraint);

            if (this.Children.Count > 0)
            {
                double width = double.IsNaN(this.Width)
                                ? this.Children.OfType<MdiWindow>().Max((w) => GetLeft(w) + w.DesiredSize.Width)
                                : this.Width;
                double height = double.IsNaN(this.Height)
                                ? this.Children.OfType<MdiWindow>().Max((w) => GetTop(w) + w.DesiredSize.Height)
                                : this.Height;
                if (double.IsNaN(width))
                {
                    width = 0d;
                }
                if (double.IsNaN(height))
                {
                    height = 0d;
                }

                return new Size(Math.Min(width, constraint.Width), Math.Min(height, constraint.Height));
            }

            return defaultValue;
        }
    }
}