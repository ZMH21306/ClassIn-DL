using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using Classin视频解析下载工具.Utils;

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
                    OnPropertyChanged(nameof(DisplayStatus));
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
                    OnPropertyChanged(nameof(Progress));
                    if (_statusFlag == DownloadStatus.Downloading)
                    {
                        OnPropertyChanged(nameof(DisplayStatus));
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
                    OnPropertyChanged(nameof(DownloadedBytes));
                    if (_statusFlag == DownloadStatus.Downloading)
                    {
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
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
                    if (_statusFlag == DownloadStatus.Downloading)
                    {
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }
        }

        private double _currentSpeedBytesPerSec;
        public double CurrentSpeedBytesPerSec
        {
            get => _currentSpeedBytesPerSec;
            set
            {
                if (Math.Abs(_currentSpeedBytesPerSec - value) > 0.01)
                {
                    _currentSpeedBytesPerSec = value;
                    OnPropertyChanged(nameof(CurrentSpeedBytesPerSec));
                    if (_statusFlag == DownloadStatus.Downloading)
                    {
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
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
                    OnPropertyChanged(nameof(RemainingTime));
                    if (_statusFlag == DownloadStatus.Downloading)
                    {
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
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

        private string _cachedDisplayStatus = string.Empty;
        private DownloadStatus _cachedStatusFlag;
        private int _cachedProgress;
        private long _cachedDownloadedBytes;
        private long _cachedTotalBytes;
        private double _cachedSpeed;
        private TimeSpan _cachedRemainingTime;

        public CancellationTokenSource? DownloadTokenSource { get; set; }

        public string StatusColorHex { get; set; } = "#FF808080";

        public string DisplayStatus
        {
            get
            {
                if (_statusFlag == DownloadStatus.Downloading &&
                    _cachedStatusFlag == _statusFlag &&
                    _cachedProgress == _progress &&
                    _cachedDownloadedBytes == _downloadedBytes &&
                    _cachedTotalBytes == _totalBytes &&
                    Math.Abs(_cachedSpeed - _currentSpeedBytesPerSec) < 0.01 &&
                    _cachedRemainingTime == _remainingTime)
                {
                    return _cachedDisplayStatus;
                }

                var newStatus = BuildDisplayStatus();
                _cachedStatusFlag = _statusFlag;
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