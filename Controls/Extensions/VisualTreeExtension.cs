using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace Hammer.MDI.Control.Extensions
{
    internal static class VisualTreeExtension
    {
        public static MdiWindow? FindMdiWindow(FrameworkElement sender)
        {
            return FindSpecificParent<MdiWindow>(sender);
        }

        public static TParent? FindSpecificParent<TParent>(FrameworkElement sender)
                   where TParent : FrameworkElement
        {
            var current = sender;
            if (current is null) return null;
            var parent = VisualTreeHelper.GetParent(current) as FrameworkElement;

            if (parent != null && parent.GetType() != typeof(TParent))
            {
                parent = FindSpecificParent<TParent>(parent);
            }

            if (parent == null && current.Parent is Popup popup && popup.Parent is FrameworkElement grandpa)
            {
                parent = FindSpecificParent<TParent>(grandpa);
            }

            if(parent is null)
            {
                return null;
            }
            else
            {
                return parent as TParent;
            }
        }
    }
}