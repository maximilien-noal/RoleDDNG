using System.Windows;
using System.Windows.Controls;

namespace Hammer.MdiControls.Panels
{
    /// <summary>
    /// Canvas is used to place child UIElements at arbitrary positions or to draw children in multiple
    /// layers.
    ///
    /// Child positions are computed from the Left, Top properties.
    /// These properties do contribute to the size of the AutoResizeCanvas.
    ///
    /// The order that children are drawn (z-order) is determined exclusively by child order.
    /// </summary>
    public class AutoResizeCanvas : Canvas
    {
        protected override Size MeasureOverride(Size constraint)
        {
            base.MeasureOverride(constraint);
            return RenderSize;
        }
    }
}