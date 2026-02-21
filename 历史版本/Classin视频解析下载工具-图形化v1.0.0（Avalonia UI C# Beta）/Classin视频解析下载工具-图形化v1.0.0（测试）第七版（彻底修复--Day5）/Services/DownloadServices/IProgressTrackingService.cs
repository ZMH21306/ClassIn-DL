using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Classin视频解析下载工具.Models;

namespace Classin视频解析下载工具.Services.DownloadServices
{
    public class ProgressInfo
    {
        public int CompletedCount { get; set; }
        public int ActiveDownloadCount { get; set; }
        public int PendingCount { get; set; }
        public int FailedCount { get; set; }
        public bool AnyActive { get; set; }
        public double OverallPercentage { get; set; }
        public long TotalBytes { get; set; }
        public long DownloadedBytes { get; set; }
        public double CurrentSpeedBytesPerSecond { get; set; }
    }

    public interface IProgressTrackingService
    {
        ProgressInfo CalculateOverallProgress(IEnumerable<VideoItem> items);
        string FormatSpeed(double bytesPerSecond);
    }
}