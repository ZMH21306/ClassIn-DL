using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;

namespace Classin视频解析下载工具.Behaviors
{
    public static class GridViewBehaviors
    {
        public static readonly AttachedProperty<bool> AutoResizeProperty =
            AvaloniaProperty.RegisterAttached<DataGrid, bool>("AutoResize", typeof(GridViewBehaviors));

        public static readonly AttachedProperty<bool> ColumnHeaderDragProperty =
            AvaloniaProperty.RegisterAttached<DataGrid, bool>("ColumnHeaderDrag", typeof(GridViewBehaviors));

        static GridViewBehaviors()
        {
            AutoResizeProperty.Changed.AddClassHandler<DataGrid>(OnAutoResizeChanged);
            ColumnHeaderDragProperty.Changed.AddClassHandler<DataGrid>(OnColumnHeaderDragChanged);
        }

        public static bool GetAutoResize(DataGrid element)
        {
            return element.GetValue(AutoResizeProperty);
        }

        public static void SetAutoResize(DataGrid element, bool value)
        {
            element.SetValue(AutoResizeProperty, value);
        }

        public static bool GetColumnHeaderDrag(DataGrid element)
        {
            return element.GetValue(ColumnHeaderDragProperty);
        }

        public static void SetColumnHeaderDrag(DataGrid element, bool value)
        {
            element.SetValue(ColumnHeaderDragProperty, value);
        }

        private static void OnAutoResizeChanged(DataGrid dataGrid, AvaloniaPropertyChangedEventArgs e)
        {
            // 在Avalonia中，DataGrid默认支持自动调整列宽
            // 这里可以根据需要添加额外的逻辑
        }

        private static void OnColumnHeaderDragChanged(DataGrid dataGrid, AvaloniaPropertyChangedEventArgs e)
        {
            // 在Avalonia中，DataGrid默认支持列标题拖拽
            // 这里可以根据需要添加额外的逻辑
        }
    }
}