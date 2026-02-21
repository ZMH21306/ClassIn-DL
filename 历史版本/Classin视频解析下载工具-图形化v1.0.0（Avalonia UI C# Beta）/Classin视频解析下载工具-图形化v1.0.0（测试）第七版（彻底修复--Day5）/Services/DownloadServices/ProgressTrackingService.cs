using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Classin视频解析下载工具.Models;

namespace Classin视频解析下载工具.Services.DownloadServices
{
    public class ProgressTrackingService : IProgressTrackingService
    {
        public ProgressInfo CalculateOverallProgress(IEnumerable<VideoItem> items)
        {
            var info = new ProgressInfo();
            var activeItems = new List<VideoItem>();

            foreach (var item in items)
            {
                if (item.StatusFlag == DownloadStatus.Completed)
                {
                    info.CompletedCount++;
                    info.TotalBytes += item.TotalBytes;
                    info.DownloadedBytes += item.TotalBytes;
                }
                else if (item.StatusFlag == DownloadStatus.Downloading)
                {
                    info.ActiveDownloadCount++;
                    activeItems.Add(item);
                    info.TotalBytes += item.TotalBytes;
                    info.DownloadedBytes += item.DownloadedBytes;
                }
                else if (item.StatusFlag == DownloadStatus.Pending)
                {
                    info.PendingCount++;
                }
                else if (item.StatusFlag == DownloadStatus.Failed)
                {
                    info.FailedCount++;
                }
            }

            info.AnyActive = info.ActiveDownloadCount > 0;

            if (info.TotalBytes > 0)
            {
                info.OverallPercentage = (double)info.DownloadedBytes / info.TotalBytes * 100;
            }

            // 计算总速度
            foreach (var item in activeItems)
            {
                info.CurrentSpeedBytesPerSecond += item.CurrentSpeedBytesPerSec;
            }

            return info;
        }

        public string FormatSpeed(double bytesPerSecond)
        {
            if (bytesPerSecond >= 1L << 30) return $"{bytesPerSecond / (1L << 30):F2} GB/s";
            if (bytesPerSecond >= 1L << 20) return $"{bytesPerSecond / (1L << 20):F2} MB/s";
            if (bytesPerSecond >= 1L << 10) return $"{bytesPerSecond / (1L << 10):F2} KB/s";
            return $"{bytesPerSecond:F2} B/s";
        }
    }
}