using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace Hammer.MDI.Control.Extensions
{
    internal static class VisualTreeExtension
    {
        public static TParent FindSpecificParent<TParent>(FrameworkElement sender)
           where TParent : FrameworkElement
        {
            var current = sender;
            if (current == null) return null;
            var parent = VisualTreeHelper.GetParent(current) as FrameworkElement;

            if (parent != null && parent.GetType() != typeof(TParent))
            {
                parent = FindSpecificParent<TParent>(parent);
            }

            if (parent == null && current.Parent is Popup)
            {
                if (((Popup)current.Parent).Parent is FrameworkElement grandpa)
                {
                    parent = FindSpecificParent<TParent>(grandpa);
                }
            }

            return parent as TParent;
        }

        public static MdiWindow FindMdiWindow(FrameworkElement sender)
        {
            return FindSpecificParent<MdiWindow>(sender);
        }
    }
}