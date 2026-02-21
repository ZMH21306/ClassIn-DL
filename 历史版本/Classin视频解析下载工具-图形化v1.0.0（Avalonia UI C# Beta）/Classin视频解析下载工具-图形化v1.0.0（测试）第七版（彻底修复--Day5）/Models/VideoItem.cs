using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using Classin视频解析下载工具.Shared.Helpers;

namespace Classin视频解析下载工具.Models
{
    public enum DownloadStatus
    {
        Pending,
        Downloading,
        Completed,
        Failed,
        Partial
    }

    public class VideoItem : INotifyPropertyChanged
    {
        private DownloadStatus _statusFlag = DownloadStatus.Pending;
        public DownloadStatus StatusFlag
        {
            get => _statusFlag;
            set
            {
                if (_statusFlag != value)
                {
                    _statusFlag = value;
                    OnPropertyChanged(nameof(StatusFlag));
                    OnPropertyChanged(nameof(IsActiveDownload));
                    OnPropertyChanged(nameof(DisplayStatus));
                }
            }
        }

        public bool IsActiveDownload => StatusFlag == DownloadStatus.Downloading;

        private int _displayIndex;
        public int DisplayIndex
        {
            get => _displayIndex;
            set
            {
                if (_displayIndex != value)
                {
                    _displayIndex = value;
                    OnPropertyChanged(nameof(DisplayIndex));
                }
            }
        }

        public string Name { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;

        private long _fileSize;
        public long FileSize
        {
            get => _fileSize;
            set
            {
                if (_fileSize != value)
                {
                    _fileSize = value;
                    OnPropertyChanged(nameof(FileSize));
                }
            }
        }

        private string _status = "等待解析";
        public string Status
        {
            get => _status;
            set
            {
                if (_status != value)
                {
                    _status = value;
                    OnPropertyChanged(nameof(Status));
                    // 只有在非下载状态下才更新DisplayStatus，下载状态下由进度更新统一处理
                    if (_statusFlag != DownloadStatus.Downloading)
                    {
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }
        }

        private int _progress;
        public int Progress
        {
            get => _progress;
            set
            {
                if (_progress != value)
                {
                    _progress = value;
                    // 只有在值变化超过1%时才通知，减少频繁更新
                    if (Math.Abs(_progress - value) >= 1 || value == 0 || value == 100)
                    {
                        OnPropertyChanged(nameof(Progress));
                    }
                }
            }
        }

        private long _downloadedBytes;
        public long DownloadedBytes
        {
            get => _downloadedBytes;
            set
            {
                if (_downloadedBytes != value)
                {
                    _downloadedBytes = value;
                    // 下载字节数变化不直接触发UI更新，由统一的进度更新处理
                }
            }
        }

        private long _totalBytes;
        public long TotalBytes
        {
            get => _totalBytes;
            set
            {
                if (_totalBytes != value)
                {
                    _totalBytes = value;
                    OnPropertyChanged(nameof(TotalBytes));
                }
            }
        }

        private double _currentSpeedBytesPerSec;
        public double CurrentSpeedBytesPerSec
        {
            get => _currentSpeedBytesPerSec;
            set
            {
                if (Math.Abs(_currentSpeedBytesPerSec - value) > 0.1) // 增加阈值，减少更新频率
                {
                    _currentSpeedBytesPerSec = value;
                    // 速度变化不直接触发UI更新，由统一的进度更新处理
                }
            }
        }

        private TimeSpan _remainingTime = TimeSpan.MaxValue;
        public TimeSpan RemainingTime
        {
            get => _remainingTime;
            set
            {
                if (_remainingTime != value)
                {
                    _remainingTime = value;
                    // 剩余时间变化不直接触发UI更新，由统一的进度更新处理
                }
            }
        }

        // 统一的进度更新方法，减少频繁的UI通知
        public void UpdateProgress(long downloaded, long total, int progress, double speed, TimeSpan remaining)
        {
            bool hasChanges = false;

            if (_downloadedBytes != downloaded)
            {
                _downloadedBytes = downloaded;
                hasChanges = true;
            }

            if (_totalBytes != total)
            {
                _totalBytes = total;
                OnPropertyChanged(nameof(TotalBytes));
                hasChanges = true;
            }

            if (_progress != progress)
            {
                _progress = progress;
                if (Math.Abs(_progress - progress) >= 1 || progress == 0 || progress == 100)
                {
                    OnPropertyChanged(nameof(Progress));
                    hasChanges = true;
                }
            }

            if (Math.Abs(_currentSpeedBytesPerSec - speed) > 0.1)
            {
                _currentSpeedBytesPerSec = speed;
                hasChanges = true;
            }

            if (_remainingTime != remaining)
            {
                _remainingTime = remaining;
                hasChanges = true;
            }

            // 只有在有变化且状态为下载中时才更新DisplayStatus
            if (hasChanges && _statusFlag == DownloadStatus.Downloading)
            {
                OnPropertyChanged(nameof(DisplayStatus));
            }
        }

        private string _phase = "等待";
        public string Phase
        {
            get => _phase;
            set
            {
                if (_phase != value)
                {
                    _phase = value;
                    OnPropertyChanged(nameof(Phase));
                    OnPropertyChanged(nameof(DisplayStatus));
                }
            }
        }

        private long _lastReportedBytes;
        public long LastReportedBytes
        {
            get => _lastReportedBytes;
            set
            {
                if (_lastReportedBytes != value)
                {
                    _lastReportedBytes = value;
                    OnPropertyChanged(nameof(LastReportedBytes));
                }
            }
        }

        private long _lastProgressUpdateTimestamp = System.Diagnostics.Stopwatch.GetTimestamp();
        public long LastProgressUpdateTimestamp
        {
            get => _lastProgressUpdateTimestamp;
            set
            {
                if (_lastProgressUpdateTimestamp != value)
                {
                    _lastProgressUpdateTimestamp = value;
                    OnPropertyChanged(nameof(LastProgressUpdateTimestamp));
                }
            }
        }

        private string _cachedDisplayStatus = string.Empty;
        private DownloadStatus _cachedStatusFlag;
        private int _cachedProgress;
        private long _cachedDownloadedBytes;
        private long _cachedTotalBytes;
        private double _cachedSpeed;
        private TimeSpan _cachedRemainingTime;
        private string _cachedStatus = string.Empty;

        public CancellationTokenSource? DownloadTokenSource { get; set; }

        private string _statusColorHex = "#FF808080";
        public string StatusColorHex
        {
            get => _statusColorHex;
            set
            {
                if (_statusColorHex != value)
                {
                    _statusColorHex = value;
                    OnPropertyChanged(nameof(StatusColorHex));
                }
            }
        }

        public string DisplayStatus
        {
            get
            {
                // 优化缓存逻辑，减少不必要的计算
                bool needUpdate = false;

                if (_statusFlag == DownloadStatus.Downloading)
                {
                    needUpdate = _cachedStatusFlag != _statusFlag ||
                                _cachedProgress != _progress ||
                                _cachedDownloadedBytes != _downloadedBytes ||
                                _cachedTotalBytes != _totalBytes ||
                                Math.Abs(_cachedSpeed - _currentSpeedBytesPerSec) > 0.1 ||
                                _cachedRemainingTime != _remainingTime;
                }
                else
                {
                    needUpdate = _cachedStatusFlag != _statusFlag ||
                                _cachedStatus != _status;
                }

                if (!needUpdate)
                {
                    return _cachedDisplayStatus;
                }

                var newStatus = BuildDisplayStatus();

                // 更新缓存
                _cachedStatusFlag = _statusFlag;
                _cachedStatus = _status;
                _cachedProgress = _progress;
                _cachedDownloadedBytes = _downloadedBytes;
                _cachedTotalBytes = _totalBytes;
                _cachedSpeed = _currentSpeedBytesPerSec;
                _cachedRemainingTime = _remainingTime;
                _cachedDisplayStatus = newStatus;

                return newStatus;
            }
        }

        private string BuildDisplayStatus()
        {
            switch (StatusFlag)
            {
                case DownloadStatus.Downloading:
                    return $"下载中 ({Progress}%) - {FormatUtils.FormatSize(DownloadedBytes)}/{FormatUtils.FormatSize(TotalBytes)} @ {FormatUtils.FormatSpeed(CurrentSpeedBytesPerSec)} - 剩余: {FormatUtils.FormatTime(RemainingTime)}";
                case DownloadStatus.Partial:
                    return $"已暂停 ({Progress}%) - {FormatUtils.FormatSize(DownloadedBytes)}/{FormatUtils.FormatSize(TotalBytes)}";
                case DownloadStatus.Completed:
                    return "下载完成";
                case DownloadStatus.Failed:
                    return "下载失败";
                default:
                    return Status;
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}