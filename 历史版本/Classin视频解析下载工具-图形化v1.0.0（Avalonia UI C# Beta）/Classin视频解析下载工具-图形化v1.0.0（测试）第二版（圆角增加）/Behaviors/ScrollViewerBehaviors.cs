using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace Classin视频解析下载工具.Behaviors
{
    public static class ScrollViewerBehaviors
    {
        public static readonly AttachedProperty<bool> AutoScrollToEndProperty =
            AvaloniaProperty.RegisterAttached<ScrollViewer, bool>("AutoScrollToEnd", typeof(ScrollViewerBehaviors));

        static ScrollViewerBehaviors()
        {
            AutoScrollToEndProperty.Changed.AddClassHandler<ScrollViewer>(OnAutoScrollToEndChanged);
        }

        public static bool GetAutoScrollToEnd(ScrollViewer element)
        {
            return element.GetValue(AutoScrollToEndProperty);
        }

        public static void SetAutoScrollToEnd(ScrollViewer element, bool value)
        {
            element.SetValue(AutoScrollToEndProperty, value);
        }

        private static void OnAutoScrollToEndChanged(ScrollViewer scrollViewer, AvaloniaPropertyChangedEventArgs e)
        {
            if (e.NewValue is bool autoScroll && autoScroll)
            {
                scrollViewer.ScrollChanged += OnScrollChanged;
            }
            else
            {
                scrollViewer.ScrollChanged -= OnScrollChanged;
            }
        }

        private static void OnScrollChanged(object? sender, ScrollChangedEventArgs e)
        {
            if (sender is ScrollViewer scrollViewer)
            {
                scrollViewer.ScrollToEnd();
            }
        }
    }
}