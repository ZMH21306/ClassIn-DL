using System;
using System.Collections.Generic;

namespace Classin视频解析下载工具.Models
{
    public class DownloadSettings
    {
        public int MaxConcurrentDownloads { get; set; } = 5;
        public int MaxDownloadThreads { get; set; } = 32;
        public string DownloadPath { get; set; } = string.Empty;
        public int BufferSizeKB { get; set; } = 1024;
        public double TimeoutHours { get; set; } = 6.0;
        public int MaxRetries { get; set; } = 3;
        public bool EnableLogging { get; set; } = true;
        public bool AutoCheckUpdates { get; set; } = true;
        public string DefaultUserAgent { get; set; } = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
        public string DefaultReferrer { get; set; } = "https://www.eeo.cn/";
        public Dictionary<string, object> CustomSettings { get; set; } = new();
    }
}