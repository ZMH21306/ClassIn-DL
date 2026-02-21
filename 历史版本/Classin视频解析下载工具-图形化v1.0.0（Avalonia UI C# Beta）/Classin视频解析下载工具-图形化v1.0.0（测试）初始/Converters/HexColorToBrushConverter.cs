using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace VideoDownloader.Converters
{
    public class HexColorToBrushConverter : IValueConverter
    {
        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is string hexColor && !string.IsNullOrEmpty(hexColor))
            {
                try
                {
                    // 移除 # 前缀如果存在
                    if (hexColor.StartsWith("#"))
                    {
                        hexColor = hexColor.Substring(1);
                    }

                    // 如果是6位十六进制颜色
                    if (hexColor.Length == 6)
                    {
                        var r = byte.Parse(hexColor.Substring(0, 2), NumberStyles.HexNumber);
                        var g = byte.Parse(hexColor.Substring(2, 2), NumberStyles.HexNumber);
                        var b = byte.Parse(hexColor.Substring(4, 2), NumberStyles.HexNumber);
                        return new SolidColorBrush(Color.FromRgb(r, g, b));
                    }
                    // 如果是8位十六进制颜色（包含Alpha）
                    else if (hexColor.Length == 8)
                    {
                        var a = byte.Parse(hexColor.Substring(0, 2), NumberStyles.HexNumber);
                        var r = byte.Parse(hexColor.Substring(2, 2), NumberStyles.HexNumber);
                        var g = byte.Parse(hexColor.Substring(4, 2), NumberStyles.HexNumber);
                        var b = byte.Parse(hexColor.Substring(6, 2), NumberStyles.HexNumber);
                        return new SolidColorBrush(Color.FromArgb(a, r, g, b));
                    }
                }
                catch
                {
                    // 解析失败时返回默认颜色
                    return new SolidColorBrush(Colors.Gray);
                }
            }
            
            return new SolidColorBrush(Colors.Gray);
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}