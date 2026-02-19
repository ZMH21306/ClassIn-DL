using System;
using System.Collections.Generic;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace VideoDownloader.Converters
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
                return Math.Max(height * 0.8, 30);
            }
            return 30.0;
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
            if (value is double actualWidth && parameter is string paramStr && double.TryParse(paramStr, out double baseSize))
            {
                var scale = actualWidth / BaseWidth;
                return Math.Max(baseSize * scale, 8);
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
}