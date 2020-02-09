using System.Windows;
using System.Windows.Media;

namespace Hammer.MDI.Control.WindowControls
{
    using System.Windows.Controls.Primitives;

    internal class WindowButton : ButtonBase
    {
        /// <summary>Identifies the <see cref="Icon"/> dependency property.</summary>
        public static readonly DependencyProperty IconProperty =
            DependencyProperty.Register(nameof(Icon), typeof(Brush), typeof(WindowButton), new UIPropertyMetadata(Brushes.Transparent));

        static WindowButton()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(WindowButton), new FrameworkPropertyMetadata(typeof(WindowButton)));
        }

        public Brush Icon
        {
            get { return (Brush)GetValue(IconProperty); }
            set { SetValue(IconProperty, value); }
        }
    }
}