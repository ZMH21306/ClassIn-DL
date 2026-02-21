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
            // 简化实现，跳过自动调整功能
            // 在Avalonia中需要不同的方法实现列宽自适应
        }

        private static void OnColumnHeaderDragChanged(ListBox listBox, AvaloniaPropertyChangedEventArgs e)
        {
            // 列头拖拽功能暂时禁用
        }

        /*private static void ListBox_SizeChanged(object? sender, SizeChangedEventArgs e)
        {
            if (sender is not ListBox listBox || listBox.View is not GridView gridView) return;

            var totalWidth = listBox.Bounds.Width - 17; // 近似滚动条宽度
            if (totalWidth <= 0) return;

            // 按照文档中定义的比例分配列宽：5%, 30%, 40%, 25%
            var ratios = new[] { 0.05, 0.3, 0.4, 0.25 };
            
            for (int i = 0; i < gridView.Columns.Count && i < ratios.Length; i++)
            {
                gridView.Columns[i].Width = totalWidth * ratios[i];
            }
        }*/
    }
}