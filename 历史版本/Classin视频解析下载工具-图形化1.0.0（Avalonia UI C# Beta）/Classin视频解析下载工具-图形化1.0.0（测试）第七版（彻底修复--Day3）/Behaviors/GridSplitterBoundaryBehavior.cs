using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Xaml.Interactivity;
using System.Linq;

namespace Classin视频解析下载工具.Behaviors
{
    /// <summary>
    /// GridSplitter边界保护行为
    /// 用于限制GridSplitter的拖拽范围，防止UI遮挡重叠和超出区域范围
    /// </summary>
    public class GridSplitterBoundaryBehavior : Behavior<GridSplitter>
    {
        /// <summary>
        /// 左侧列的最小宽度
        /// </summary>
        public static readonly StyledProperty<double> LeftColumnMinWidthProperty = 
            AvaloniaProperty.Register<GridSplitterBoundaryBehavior, double>(nameof(LeftColumnMinWidth), 150);

        /// <summary>
        /// 右侧列的最小宽度
        /// </summary>
        public static readonly StyledProperty<double> RightColumnMinWidthProperty = 
            AvaloniaProperty.Register<GridSplitterBoundaryBehavior, double>(nameof(RightColumnMinWidth), 300);

        /// <summary>
        /// 左侧列的最小宽度
        /// </summary>
        public double LeftColumnMinWidth
        {
            get => GetValue(LeftColumnMinWidthProperty);
            set => SetValue(LeftColumnMinWidthProperty, value);
        }

        /// <summary>
        /// 右侧列的最小宽度
        /// </summary>
        public double RightColumnMinWidth
        {
            get => GetValue(RightColumnMinWidthProperty);
            set => SetValue(RightColumnMinWidthProperty, value);
        }

        private Grid _parentGrid;
        private int _splitterColumnIndex;
        private double _gridTotalWidth;
        private bool _isDragging;

        protected override void OnAttached()
        {
            base.OnAttached();
            
            if (AssociatedObject != null)
            {
                AssociatedObject.DragStarted += OnDragStarted;
                AssociatedObject.DragDelta += OnDragDelta;
                AssociatedObject.DragCompleted += OnDragCompleted;
                
                // 初始化时获取父Grid和列索引
                InitializeParentGridInfo();
            }
        }

        protected override void OnDetaching()
        {
            base.OnDetaching();
            
            if (AssociatedObject != null)
            {
                AssociatedObject.DragStarted -= OnDragStarted;
                AssociatedObject.DragDelta -= OnDragDelta;
                AssociatedObject.DragCompleted -= OnDragCompleted;
            }
        }

        /// <summary>
        /// 初始化父Grid信息
        /// </summary>
        private void InitializeParentGridInfo()
        {
            if (AssociatedObject?.Parent is Grid grid)
            {
                _parentGrid = grid;
                
                // 查找GridSplitter所在的列索引
                for (int i = 0; i < grid.ColumnDefinitions.Count; i++)
                {
                    var element = grid.Children.FirstOrDefault(child => 
                        Grid.GetColumn(child) == i && child == AssociatedObject);
                    
                    if (element != null)
                    {
                        _splitterColumnIndex = i;
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// 拖拽开始事件处理
        /// </summary>
        private void OnDragStarted(object sender, VectorEventArgs e)
        {
            _isDragging = true;
            
            // 重新计算Grid总宽度，确保使用最新值
            if (_parentGrid != null)
            {
                _gridTotalWidth = _parentGrid.Bounds.Width;
            }
        }

        /// <summary>
        /// 拖拽过程事件处理
        /// </summary>
        private void OnDragDelta(object sender, VectorEventArgs e)
        {
            if (!_isDragging || _parentGrid == null || _splitterColumnIndex <= 0 || _splitterColumnIndex >= _parentGrid.ColumnDefinitions.Count - 1)
                return;

            // 计算当前拖拽位置
            double horizontalChange = e.Vector.X;
            
            // 获取左侧和右侧列的当前宽度
            var leftColumn = _parentGrid.ColumnDefinitions[_splitterColumnIndex - 1];
            var rightColumn = _parentGrid.ColumnDefinitions[_splitterColumnIndex + 1];
            
            // 计算新的宽度
            double newLeftWidth = leftColumn.Width.Value + horizontalChange;
            double newRightWidth = rightColumn.Width.Value - horizontalChange;
            
            // 检查边界限制
            if (newLeftWidth >= LeftColumnMinWidth && newRightWidth >= RightColumnMinWidth)
            {
                // 更新列宽度
                leftColumn.Width = new GridLength(newLeftWidth, GridUnitType.Pixel);
                rightColumn.Width = new GridLength(newRightWidth, GridUnitType.Pixel);
            }
        }

        /// <summary>
        /// 拖拽完成事件处理
        /// </summary>
        private void OnDragCompleted(object sender, VectorEventArgs e)
        {
            _isDragging = false;
        }
    }
}