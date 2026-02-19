using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;

namespace Classin视频解析下载工具.Behaviors
{
    /// <summary>
    /// GridView 相关行为
    /// </summary>
    public static class GridViewBehaviors
    {
        public static readonly AttachedProperty<bool> AutoResizeProperty = AvaloniaProperty.RegisterAttached<ListBox, bool>("AutoResize", typeof(GridViewBehaviors));
        public static readonly AttachedProperty<bool> ColumnHeaderDragProperty = AvaloniaProperty.RegisterAttached<ListBox, bool>("ColumnHeaderDrag", typeof(GridViewBehaviors));

        static GridViewBehaviors()
        {
            AutoResizeProperty.Changed.AddClassHandler<ListBox>(OnAutoResizeChanged);
            ColumnHeaderDragProperty.Changed.AddClassHandler<ListBox>(OnColumnHeaderDragChanged);
        }

        public static bool GetAutoResize(ListBox element) => element.GetValue(AutoResizeProperty);
        public static void SetAutoResize(ListBox element, bool value) => element.SetValue(AutoResizeProperty, value);

        public static bool GetColumnHeaderDrag(ListBox element) => element.GetValue(ColumnHeaderDragProperty);
        public static void SetColumnHeaderDrag(ListBox element, bool value) => element.SetValue(ColumnHeaderDragProperty, value);

        private static void OnAutoResizeChanged(ListBox listBox, AvaloniaPropertyChangedEventArgs e)
        {
            // 暂时禁用自动调整功能，使用固定列宽布局
            // 后续可以在此基础上添加更完善的自适应逻辑
        }

        private static void OnColumnHeaderDragChanged(ListBox listBox, AvaloniaPropertyChangedEventArgs e)
        {
            // 列头拖拽功能暂时禁用
        }
    }

    /// <summary>
    /// ListView 相关行为
    /// </summary>
    public static class ListViewBehaviors
    {
        public static readonly AttachedProperty<double> RowHeightProperty = AvaloniaProperty.RegisterAttached<ListBox, double>("RowHeight", typeof(ListViewBehaviors));
        public static readonly AttachedProperty<bool> ColumnHeaderDragProperty = AvaloniaProperty.RegisterAttached<ListBox, bool>("ColumnHeaderDrag", typeof(ListViewBehaviors));

        static ListViewBehaviors()
        {
            RowHeightProperty.Changed.AddClassHandler<ListBox>(OnRowHeightChanged);
            ColumnHeaderDragProperty.Changed.AddClassHandler<ListBox>(OnColumnHeaderDragChanged);
        }

        public static double GetRowHeight(ListBox element) => element.GetValue(RowHeightProperty);
        public static void SetRowHeight(ListBox element, double value) => element.SetValue(RowHeightProperty, value);

        public static bool GetColumnHeaderDrag(ListBox element) => element.GetValue(ColumnHeaderDragProperty);
        public static void SetColumnHeaderDrag(ListBox element, bool value) => element.SetValue(ColumnHeaderDragProperty, value);

        private static void OnRowHeightChanged(ListBox listBox, AvaloniaPropertyChangedEventArgs e)
        {
            // 行高功能暂时使用固定值，后续可以在此基础上完善动态行高逻辑
        }

        private static void OnColumnHeaderDragChanged(ListBox listBox, AvaloniaPropertyChangedEventArgs e)
        {
            // 列头拖拽功能暂时禁用
        }
    }

    /// <summary>
    /// ScrollViewer 相关行为
    /// </summary>
    public static class ScrollViewerBehaviors
    {
        public static readonly AttachedProperty<bool> AutoScrollToEndProperty = AvaloniaProperty.RegisterAttached<ScrollViewer, bool>("AutoScrollToEnd", typeof(ScrollViewerBehaviors));
        private static readonly AvaloniaProperty<bool> ShouldAutoScrollProperty = AvaloniaProperty.RegisterAttached<ScrollViewer, bool>("ShouldAutoScroll", typeof(ScrollViewerBehaviors), true);

        static ScrollViewerBehaviors()
        {
            AutoScrollToEndProperty.Changed.AddClassHandler<ScrollViewer>(OnAutoScrollToEndChanged);
        }

        public static bool GetAutoScrollToEnd(ScrollViewer element) => element.GetValue(AutoScrollToEndProperty);
        public static void SetAutoScrollToEnd(ScrollViewer element, bool value) => element.SetValue(AutoScrollToEndProperty, value);

        private static bool GetShouldAutoScroll(ScrollViewer element) => element.GetValue(ShouldAutoScrollProperty) is bool value && value;
        private static void SetShouldAutoScroll(ScrollViewer element, bool value) => element.SetValue(ShouldAutoScrollProperty, value);

        private static void OnAutoScrollToEndChanged(ScrollViewer scrollViewer, AvaloniaPropertyChangedEventArgs e)
        {
            if (e.NewValue is bool autoScroll && autoScroll)
            {
                scrollViewer.ScrollChanged += OnScrollChanged;
                scrollViewer.ScrollToEnd();
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
                if (e.ExtentDelta.Y > 0 || e.ExtentDelta.X > 0)
                {
                    if (GetShouldAutoScroll(scrollViewer))
                    {
                        scrollViewer.ScrollToEnd();
                    }
                }
                else if (e.OffsetDelta.Y != 0 || e.OffsetDelta.X != 0)
                {
                    var verticalOffset = scrollViewer.Offset.Y;
                    var viewportHeight = scrollViewer.Viewport.Height;
                    var extentHeight = scrollViewer.Extent.Height;

                    if (verticalOffset + viewportHeight >= extentHeight - 5)
                    {
                        SetShouldAutoScroll(scrollViewer, true);
                    }
                    else
                    {
                        SetShouldAutoScroll(scrollViewer, false);
                    }
                }
            }
        }
    }

    /// <summary>
    /// Window 相关行为
    /// </summary>
    public static class WindowBehaviors
    {
        public static readonly AttachedProperty<bool> DisableCloseProperty = AvaloniaProperty.RegisterAttached<Window, bool>("DisableClose", typeof(WindowBehaviors));

        static WindowBehaviors()
        {
            DisableCloseProperty.Changed.AddClassHandler<Window>(OnDisableCloseChanged);
        }

        public static bool GetDisableClose(Window element) => element.GetValue(DisableCloseProperty);
        public static void SetDisableClose(Window element, bool value) => element.SetValue(DisableCloseProperty, value);

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