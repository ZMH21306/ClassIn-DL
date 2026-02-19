using System;
using System.Text;

namespace Classin视频解析下载工具.Shared.Helpers
{
    public static class FormatUtils
    {
        public static string FormatSize(long bytes)
        {
            if (bytes < 0) return "0 B";

            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            double len = bytes;
            int order = 0;

            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len /= 1024;
            }

            return $"{len:0.##} {sizes[order]}";
        }

        public static string FormatSpeed(double bytesPerSecond)
        {
            if (bytesPerSecond < 0) return "0 B/s";

            string[] sizes = { "B/s", "KB/s", "MB/s", "GB/s" };
            double speed = bytesPerSecond;
            int order = 0;

            while (speed >= 1024 && order < sizes.Length - 1)
            {
                order++;
                speed /= 1024;
            }

            return $"{speed:0.##} {sizes[order]}";
        }

        public static string FormatTime(TimeSpan timeSpan)
        {
            if (timeSpan == TimeSpan.MaxValue || timeSpan.TotalSeconds <= 0)
                return "未知";

            if (timeSpan.TotalDays >= 1)
                return $"{timeSpan.Days}天{timeSpan.Hours}小时";

            if (timeSpan.TotalHours >= 1)
                return $"{timeSpan.Hours}小时{timeSpan.Minutes}分钟";

            if (timeSpan.TotalMinutes >= 1)
                return $"{timeSpan.Minutes}分钟{timeSpan.Seconds}秒";

            return $"{timeSpan.Seconds}秒";
        }

        public static string FixEncoding(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                byte[] bytes = Encoding.GetEncoding("GB18030").GetBytes(text);
                return Encoding.UTF8.GetString(bytes);
            }
            catch
            {
                return text;
            }
        }
    }
}