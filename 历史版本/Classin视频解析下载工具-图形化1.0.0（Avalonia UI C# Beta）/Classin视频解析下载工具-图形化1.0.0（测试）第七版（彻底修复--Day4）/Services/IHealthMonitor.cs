using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Classin视频解析下载工具.Services
{
    public class HealthStats
    {
        public int InitializationCount { get; set; }
        public int ConfigurationChanges { get; set; }
        public int TotalDownloadsAttempted { get; set; }
        public int SuccessfulDownloads { get; set; }
        public int FailedDownloads { get; set; }
        public int ActiveDownloads { get; set; }
        public long TotalDownloadedBytes { get; set; }
        public DateTime? LastInitializationTime { get; set; }
        public DateTime? LastConfigurationChange { get; set; }
        public DateTime? LastSuccessfulDownload { get; set; }
        public DateTime? LastFailedDownload { get; set; }

        public double SuccessRate => TotalDownloadsAttempted > 0 
            ? (double)SuccessfulDownloads / TotalDownloadsAttempted * 100 
            : 0;
    }

    public interface IHealthMonitor : IDisposable
    {
        void RecordInitialization();
        void RecordConfigurationChange(string settingName, string newValue);
        void RecordDownloadStart(string videoName);
        void RecordDownloadComplete(string videoName, long fileSize);
        void RecordDownloadFailure(string videoName, string error);
        void Reset();
        HealthStats GetStats();
    }
}