using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;

namespace VideoDownloader.Behaviors
{
    public static class WindowBehaviors
    {
        public static readonly AttachedProperty<bool> DisableCloseProperty =
            AvaloniaProperty.RegisterAttached<Window, bool>("DisableClose", typeof(WindowBehaviors));

        static WindowBehaviors()
        {
            DisableCloseProperty.Changed.AddClassHandler<Window>(OnDisableCloseChanged);
        }

        public static bool GetDisableClose(Window element)
        {
            return element.GetValue(DisableCloseProperty);
        }

        public static void SetDisableClose(Window element, bool value)
        {
            element.SetValue(DisableCloseProperty, value);
        }

        private static void OnDisableCloseChanged(Window window, AvaloniaPropertyChangedEventArgs e)
        {
            if (e.NewValue is bool disableClose && disableClose)
            {
                window.Closing += OnWindowClosing;
            }
            else
            {
                window.Closing -= OnWindowClosing;
            }
        }

        private static void OnWindowClosing(object? sender, WindowClosingEventArgs e)
        {
            e.Cancel = true;
        }
    }
}