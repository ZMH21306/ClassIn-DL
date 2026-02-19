using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;

namespace Classin视频解析下载工具.Behaviors
{
    public static class GridViewBehaviors
    {
        public static readonly AttachedProperty<bool> AutoResizeProperty =
            AvaloniaProperty.RegisterAttached<ListBox, bool>("AutoResize", typeof(GridViewBehaviors));

        public static readonly AttachedProperty<bool> ColumnHeaderDragProperty =
            AvaloniaProperty.RegisterAttached<ListBox, bool>("ColumnHeaderDrag", typeof(GridViewBehaviors));

        static GridViewBehaviors()
        {
            AutoResizeProperty.Changed.AddClassHandler<ListBox>(OnAutoResizeChanged);
            ColumnHeaderDragProperty.Changed.AddClassHandler<ListBox>(OnColumnHeaderDragChanged);
        }

        public static bool GetAutoResize(ListBox element)
        {
            return element.GetValue(AutoResizeProperty);
        }

        public static void SetAutoResize(ListBox element, bool value)
        {
            element.SetValue(AutoResizeProperty, value);
        }

        public static bool GetColumnHeaderDrag(ListBox element)
        {
            return element.GetValue(ColumnHeaderDragProperty);
        }

        public static void SetColumnHeaderDrag(ListBox element, bool value)
        {
            element.SetValue(ColumnHeaderDragProperty, value);
        }

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
}