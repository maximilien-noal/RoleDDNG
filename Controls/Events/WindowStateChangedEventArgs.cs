using System.Windows;

namespace Hammer.MDI.Control.Events
{
    public sealed class WindowStateChangedEventArgs : RoutedEventArgs
    {
        public WindowStateChangedEventArgs(RoutedEvent routedEvent, WindowState oldValue, WindowState newValue)
           : base(routedEvent)
        {
            NewValue = newValue;
            OldValue = oldValue;
        }

        public WindowState NewValue { get; private set; }

        public WindowState OldValue { get; private set; }
    }
}