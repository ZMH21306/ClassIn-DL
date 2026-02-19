using System;
using System.Collections.Generic;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace Classin视频解析下载工具.Converters
{
    public class DarkenColorConverter : IValueConverter
    {
        public double DarkenFactor { get; set; } = 0.3;

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is IBrush brush)
            {
                var color = brush is SolidColorBrush solidBrush ? solidBrush.Color : Colors.Black;
                var r = (byte)(color.R * (1 - DarkenFactor));
                var g = (byte)(color.G * (1 - DarkenFactor));
                var b = (byte)(color.B * (1 - DarkenFactor));
                return new SolidColorBrush(Color.FromRgb(r, g, b));
            }
            return value;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class HeightConverter : IValueConverter
    {
        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is double height)
            {
                double itemHeight = (height - 5) / 10.0;
                return Math.Clamp(itemHeight, 30, 80);
            }
            return 40.0;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class IndexFormatConverter : IValueConverter
    {
        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is int index)
            {
                return $"{index:D2}";
            }
            return "00";
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class ScalingConverter : IValueConverter
    {
        public double BaseWidth { get; set; } = 1440;

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is double width && parameter is string paramString && double.TryParse(paramString, out double baseSize))
            {
                double scale = Math.Clamp(width / BaseWidth, 0.7, 1.3);
                return baseSize * scale;
            }
            return parameter ?? 12.0;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class PathMaxWidthConverter : IMultiValueConverter
    {
        public object? Convert(IList<object?> values, Type targetType, object? parameter, CultureInfo culture)
        {
            if (values.Count >= 3 &&
                values[0] is double gridWidth &&
                values[1] is double pathWidth &&
                values[2] is double buttonWidth)
            {
                var availableWidth = gridWidth - 300 - buttonWidth - 20;
                return Math.Max(availableWidth, 100);
            }
            return 200.0;
        }
    }

    /// <summary>
    /// 行高转换器 - 根据容器高度动态计算行高
    /// </summary>
    public class RowHeightConverter : IValueConverter
    {
        /// <summary>
        /// 基础行数，用于计算行高
        /// </summary>
        public int BaseRowCount { get; set; } = 8;
        
        /// <summary>
        /// 最小行高
        /// </summary>
        public double MinHeight { get; set; } = 28.0;
        
        /// <summary>
        /// 最大行高
        /// </summary>
        public double MaxHeight { get; set; } = 45.0;
        
        /// <summary>
        /// 表头高度
        /// </summary>
        public double HeaderHeight { get; set; } = 35.0;
        
        /// <summary>
        /// 边距
        /// </summary>
        public double Margin { get; set; } = 8.0;

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is double containerHeight && containerHeight > 0)
            {
                // 计算可用于行显示的高度
                double availableHeight = containerHeight - HeaderHeight - Margin;
                
                if (availableHeight <= 0)
                    return MinHeight;
                
                // 计算理想行高
                double idealHeight = availableHeight / BaseRowCount;
                
                // 限制在合理范围内
                double clampedHeight = Math.Clamp(idealHeight, MinHeight, MaxHeight);
                
                // 返回整数值以确保像素对齐
                return Math.Round(clampedHeight);
            }
            
            return 32.0; // 默认行高
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// 列宽转换器 - 根据容器宽度动态计算各列宽度
    /// </summary>
    public class ColumnWidthConverter : IMultiValueConverter
    {
        /// <summary>
        /// 列类型枚举
        /// </summary>
        public enum ColumnType
        {
            Index,      // 序号列
            CourseName, // 课程名称列
            Status,     // 状态列
            Actions     // 操作列
        }
        
        /// <summary>
        /// 滚动条预估宽度
        /// </summary>
        public double ScrollBarWidth { get; set; } = 17.0;
        
        /// <summary>
        /// 边距
        /// </summary>
        public double Margin { get; set; } = 10.0;

        public object? Convert(IList<object?> values, Type targetType, object? parameter, CultureInfo culture)
        {
            string? columnTypeStr = null;
            
            if (values.Count < 2 ||
                !(values[0] is double containerWidth) ||
                !(values[1] is string tempColumnTypeStr))
            {
                return GetDefaultWidth(columnTypeStr);
            }
            
            columnTypeStr = tempColumnTypeStr;

            // 解析列类型
            if (!Enum.TryParse<ColumnType>(columnTypeStr, out var columnType))
            {
                return GetDefaultWidth(columnTypeStr);
            }

            // 计算总可用宽度
            double totalAvailableWidth = containerWidth - ScrollBarWidth - Margin;
            
            if (totalAvailableWidth <= 0)
                return GetDefaultWidth(columnTypeStr);

            // 根据列类型返回相应宽度
            return CalculateColumnWidth(columnType, totalAvailableWidth);
        }

        private double CalculateColumnWidth(ColumnType columnType, double totalWidth)
        {
            return columnType switch
            {
                ColumnType.Index => 50.0,           // 序号列固定50px
                ColumnType.Status => 150.0,          // 状态列固定150px
                ColumnType.Actions => 180.0,         // 操作列固定180px
                ColumnType.CourseName => Math.Max(150.0, totalWidth - 50.0 - 150.0 - 180.0), // 剩余空间
                _ => 100.0
            };
        }

        private double GetDefaultWidth(string? columnTypeStr)
        {
            return columnTypeStr?.ToLower() switch
            {
                "index" => 50.0,
                "status" => 150.0,
                "actions" => 180.0,
                "coursename" => 200.0,
                _ => 100.0
            };
        }

        public object? ConvertBack(IList<object?> values, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// 增强的缩放转换器 - 支持更精确的缩放控制
    /// </summary>
    public class EnhancedScalingConverter : IValueConverter
    {
        /// <summary>
        /// 基准宽度
        /// </summary>
        public double BaseWidth { get; set; } = 1440.0;
        
        /// <summary>
        /// 最小缩放比例
        /// </summary>
        public double MinScale { get; set; } = 0.7;
        
        /// <summary>
        /// 最大缩放比例
        /// </summary>
        public double MaxScale { get; set; } = 1.3;
        
        /// <summary>
        /// 是否启用平滑缩放
        /// </summary>
        public bool SmoothScaling { get; set; } = true;

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is double width && parameter is string paramString && 
                double.TryParse(paramString, out double baseSize))
            {
                // 计算缩放比例
                double scale = width / BaseWidth;
                
                // 应用缩放限制
                scale = Math.Clamp(scale, MinScale, MaxScale);
                
                // 如果启用平滑缩放，在边界处应用缓动效果
                if (SmoothScaling)
                {
                    scale = ApplySmoothing(scale);
                }
                
                double result = baseSize * scale;
                
                // 对于字体大小，确保为整数
                if (targetType == typeof(double) || targetType == typeof(int))
                {
                    return Math.Round(result);
                }
                
                return result;
            }
            
            return parameter ?? 12.0;
        }

        private double ApplySmoothing(double scale)
        {
            // 在边界附近应用平滑过渡
            const double smoothRange = 0.1;
            
            if (scale < MinScale + smoothRange)
            {
                double t = (scale - MinScale) / smoothRange;
                return MinScale + (scale - MinScale) * SmoothStep(t);
            }
            
            if (scale > MaxScale - smoothRange)
            {
                double t = (MaxScale - scale) / smoothRange;
                return MaxScale - (MaxScale - scale) * SmoothStep(t);
            }
            
            return scale;
        }

        private double SmoothStep(double t)
        {
            // 平滑插值函数
            return t * t * (3 - 2 * t);
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}