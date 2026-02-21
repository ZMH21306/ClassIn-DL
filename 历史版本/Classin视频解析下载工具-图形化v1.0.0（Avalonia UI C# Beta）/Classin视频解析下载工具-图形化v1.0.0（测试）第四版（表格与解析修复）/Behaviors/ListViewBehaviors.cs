using System;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.VisualTree;

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
            // 行高通过绑定和转换器在XAML中处理
        }

        private static void OnColumnHeaderDragChanged(ListBox listBox, AvaloniaPropertyChangedEventArgs e)
        {
            if (e.NewValue is bool enableDrag && enableDrag)
            {
                listBox.AttachedToVisualTree += ListBox_AttachedToVisualTree;
            }
            else
            {
                listBox.AttachedToVisualTree -= ListBox_AttachedToVisualTree;
            }
        }

        private static void ListBox_AttachedToVisualTree(object? sender, VisualTreeAttachmentEventArgs e)
        {
            if (sender is ListBox listBox)
            {
                // 延迟执行以确保模板已应用
                Avalonia.Threading.Dispatcher.UIThread.Post(() => EnableColumnHeaderDrag(listBox), 
                    Avalonia.Threading.DispatcherPriority.Loaded);
            }
        }

        private static void EnableColumnHeaderDrag(ListBox listBox)
        {
            // 简化实现，跳过列头拖拽功能
            // 在Avalonia中GridView的支持有限
        }

        /*private static GridViewColumnHeader? FindColumnHeader(ListBox listBox, GridViewColumn column)
        {
            // 查找对应列的列头控件
            var columnHeaders = listBox.GetLogicalDescendants()
                .OfType<GridViewColumnHeader>()
                .Where(header => header.Column == column)
                .ToList();

            return columnHeaders.FirstOrDefault();
        }*/

        /*private static void ColumnHeader_Loaded(object? sender, RoutedEventArgs e)
        {
            if (sender is not GridViewColumnHeader header) return;

            var thumb = FindVisualChild<Thumb>(header);
            if (thumb == null) return;

            thumb.DragDelta += (s, args) =>
            {
                if (header.Column != null && header.Column.Width >= 20)
                {
                    header.Column.Width = Math.Max(20, header.Column.Width + args.HorizontalChange);
                }
            };
        }*/

        /*private static T? FindVisualChild<T>(IVisual parent) where T : class
        {
            if (parent is T result)
                return result;

            foreach (var child in parent.VisualChildren)
            {
                var found = FindVisualChild<T>(child);
                if (found != null)
                    return found;
            }

            return null;
        }*/
    }
}