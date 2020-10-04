namespace RoleDDNG.Role.UserControls.Decorators
{
    using System;
    using System.Windows;
    using System.Windows.Controls;

    public class RestrictWidthDecorator : Decorator
    {
        private Size _lastArrangeSize = new Size(double.PositiveInfinity, double.PositiveInfinity);

        protected override Size MeasureOverride(Size constraint)
        {
            base.MeasureOverride(new Size(Math.Min(this._lastArrangeSize.Width, constraint.Width), constraint.Height));
            return new Size(0, Child.DesiredSize.Height);
        }

        protected override Size ArrangeOverride(Size arrangeSize)
        {
            if (_lastArrangeSize != arrangeSize)
            {
                _lastArrangeSize = arrangeSize;
                base.MeasureOverride(arrangeSize);
            }

            return base.ArrangeOverride(arrangeSize);
        }
    }
}