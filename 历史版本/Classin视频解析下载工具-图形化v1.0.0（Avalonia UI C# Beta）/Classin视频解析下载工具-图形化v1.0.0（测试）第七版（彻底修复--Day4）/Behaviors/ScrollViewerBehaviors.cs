using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;

namespace Classin视频解析下载工具.Behaviors
{
    public static class ScrollViewerBehaviors
    {
        public static readonly AttachedProperty<bool> AutoScrollToEndProperty =
            AvaloniaProperty.RegisterAttached<ScrollViewer, bool>("AutoScrollToEnd", typeof(ScrollViewerBehaviors));

        // 跟踪是否应该自动滚动
        private static readonly AvaloniaProperty<bool> ShouldAutoScrollProperty =
            AvaloniaProperty.RegisterAttached<ScrollViewer, bool>("ShouldAutoScroll", typeof(ScrollViewerBehaviors), true);

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

        private static bool GetShouldAutoScroll(ScrollViewer element)
        {
            return (bool)element.GetValue(ShouldAutoScrollProperty);
        }

        private static void SetShouldAutoScroll(ScrollViewer element, bool value)
        {
            element.SetValue(ShouldAutoScrollProperty, value);
        }

        private static void OnAutoScrollToEndChanged(ScrollViewer scrollViewer, AvaloniaPropertyChangedEventArgs e)
        {
            if (e.NewValue is bool autoScroll && autoScroll)
            {
                scrollViewer.ScrollChanged += OnScrollChanged;
                // 初始加载时自动滚动到底部
                scrollViewer.ScrollToEnd();
                // 初始状态下应该自动滚动
                SetShouldAutoScroll(scrollViewer, true);
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
                // 检查是否是由内容变化引起的滚动
                if (e.ExtentDelta.Y > 0 || e.ExtentDelta.X > 0)
                {
                    // 内容增加了，检查是否应该自动滚动
                    if (GetShouldAutoScroll(scrollViewer))
                    {
                        scrollViewer.ScrollToEnd();
                    }
                }
                else if (e.OffsetDelta.Y != 0 || e.OffsetDelta.X != 0)
                {
                    // 用户手动滚动了，检查是否在底部
                    var verticalOffset = scrollViewer.Offset.Y;
                    var viewportHeight = scrollViewer.Viewport.Height;
                    var extentHeight = scrollViewer.Extent.Height;

                    // 如果用户滚动到底部附近，恢复自动滚动
                    if (verticalOffset + viewportHeight >= extentHeight - 5)
                    {
                        SetShouldAutoScroll(scrollViewer, true);
                    }
                    else
                    {
                        // 用户滚动离开底部，暂停自动滚动
                        SetShouldAutoScroll(scrollViewer, false);
                    }
                }
            }
        }
    }
}