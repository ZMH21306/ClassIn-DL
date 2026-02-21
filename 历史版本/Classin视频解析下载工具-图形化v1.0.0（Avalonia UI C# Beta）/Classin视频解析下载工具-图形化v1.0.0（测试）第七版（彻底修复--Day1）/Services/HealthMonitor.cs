using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Classin视频解析下载工具.Services
{
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

    public class HealthMonitor : IHealthMonitor
    {
        private readonly object _lockObject = new();
        private bool _disposed;
        
        private HealthStats _stats = new();

        public void RecordInitialization()
        {
            lock (_lockObject)
            {
                _stats.InitializationCount++;
                _stats.LastInitializationTime = DateTime.Now;
            }
        }

        public void RecordConfigurationChange(string settingName, string newValue)
        {
            lock (_lockObject)
            {
                _stats.ConfigurationChanges++;
                _stats.LastConfigurationChange = DateTime.Now;
            }
        }

        public void RecordDownloadStart(string videoName)
        {
            lock (_lockObject)
            {
                _stats.TotalDownloadsAttempted++;
                _stats.ActiveDownloads++;
            }
        }

        public void RecordDownloadComplete(string videoName, long fileSize)
        {
            lock (_lockObject)
            {
                _stats.SuccessfulDownloads++;
                _stats.TotalDownloadedBytes += fileSize;
                _stats.ActiveDownloads = Math.Max(0, _stats.ActiveDownloads - 1);
                _stats.LastSuccessfulDownload = DateTime.Now;
            }
        }

        public void RecordDownloadFailure(string videoName, string error)
        {
            lock (_lockObject)
            {
                _stats.FailedDownloads++;
                _stats.ActiveDownloads = Math.Max(0, _stats.ActiveDownloads - 1);
                _stats.LastFailedDownload = DateTime.Now;
            }
        }

        public void Reset()
        {
            lock (_lockObject)
            {
                _stats = new HealthStats();
            }
        }

        public HealthStats GetStats()
        {
            lock (_lockObject)
            {
                return new HealthStats
                {
                    InitializationCount = _stats.InitializationCount,
                    ConfigurationChanges = _stats.ConfigurationChanges,
                    TotalDownloadsAttempted = _stats.TotalDownloadsAttempted,
                    SuccessfulDownloads = _stats.SuccessfulDownloads,
                    FailedDownloads = _stats.FailedDownloads,
                    ActiveDownloads = _stats.ActiveDownloads,
                    TotalDownloadedBytes = _stats.TotalDownloadedBytes,
                    LastInitializationTime = _stats.LastInitializationTime,
                    LastConfigurationChange = _stats.LastConfigurationChange,
                    LastSuccessfulDownload = _stats.LastSuccessfulDownload,
                    LastFailedDownload = _stats.LastFailedDownload
                };
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
        }
    }

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
}