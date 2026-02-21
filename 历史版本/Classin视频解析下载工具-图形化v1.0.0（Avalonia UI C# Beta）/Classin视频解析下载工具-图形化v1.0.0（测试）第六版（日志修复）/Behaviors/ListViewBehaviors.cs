using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;

namespace Classin视频解析下载工具.Behaviors
{
    public static class ListViewBehaviors
    {
        public static readonly AttachedProperty<double> RowHeightProperty =
            AvaloniaProperty.RegisterAttached<ListBox, double>("RowHeight", typeof(ListViewBehaviors));

        public static readonly AttachedProperty<bool> ColumnHeaderDragProperty =
            AvaloniaProperty.RegisterAttached<ListBox, bool>("ColumnHeaderDrag", typeof(ListViewBehaviors));

        static ListViewBehaviors()
        {
            RowHeightProperty.Changed.AddClassHandler<ListBox>(OnRowHeightChanged);
            ColumnHeaderDragProperty.Changed.AddClassHandler<ListBox>(OnColumnHeaderDragChanged);
        }

        public static double GetRowHeight(ListBox element)
        {
            return element.GetValue(RowHeightProperty);
        }

        public static void SetRowHeight(ListBox element, double value)
        {
            element.SetValue(RowHeightProperty, value);
        }

        public static bool GetColumnHeaderDrag(ListBox element)
        {
            return element.GetValue(ColumnHeaderDragProperty);
        }

        public static void SetColumnHeaderDrag(ListBox element, bool value)
        {
            element.SetValue(ColumnHeaderDragProperty, value);
        }

        private static void OnRowHeightChanged(ListBox listBox, AvaloniaPropertyChangedEventArgs e)
        {
            // 行高功能暂时使用固定值，后续可以在此基础上完善动态行高逻辑
        }

        private static void OnColumnHeaderDragChanged(ListBox listBox, AvaloniaPropertyChangedEventArgs e)
        {
            // 列头拖拽功能暂时禁用
        }
    }
}