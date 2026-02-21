using System;

namespace Classin视频解析下载工具.Models
{
    public class ApplicationState
    {
        public bool IsInitialized { get; set; }
        public DateTime LastActivityTime { get; set; } = DateTime.Now;
        public int TotalDownloads { get; set; }
        public int CompletedDownloads { get; set; }
        public int FailedDownloads { get; set; }
        public long TotalDownloadedBytes { get; set; }
        public string CurrentOperation { get; set; } = string.Empty;
        public bool IsNetworkAvailable { get; set; } = true;
        public string LastErrorMessage { get; set; } = string.Empty;
        public DateTime LastErrorTime { get; set; } = DateTime.MinValue;
    }
}