using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;

namespace VideoDownloader.Behaviors
{
    public static class ListViewBehaviors
    {
        public static readonly AttachedProperty<double> RowHeightProperty =
            AvaloniaProperty.RegisterAttached<ListBox, double>("RowHeight", typeof(ListViewBehaviors));

        static ListViewBehaviors()
        {
            RowHeightProperty.Changed.AddClassHandler<ListBox>(OnRowHeightChanged);
        }

        public static double GetRowHeight(ListBox element)
        {
            return element.GetValue(RowHeightProperty);
        }

        public static void SetRowHeight(ListBox element, double value)
        {
            element.SetValue(RowHeightProperty, value);
        }

        private static void OnRowHeightChanged(ListBox listBox, AvaloniaPropertyChangedEventArgs e)
        {
            // 在Avalonia中，我们可以通过样式或数据模板来控制行高
            // 这里只是一个占位符实现
        }
    }
}