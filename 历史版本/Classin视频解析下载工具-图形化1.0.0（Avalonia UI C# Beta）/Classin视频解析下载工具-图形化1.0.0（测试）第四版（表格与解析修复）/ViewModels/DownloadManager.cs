using System;
using System.Collections.ObjectModel;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using Classin视频解析下载工具.Models;
using Classin视频解析下载工具.Services;
using Classin视频解析下载工具.Utils;

namespace Classin视频解析下载工具.ViewModels
{
    public sealed class DownloadManager : INotifyPropertyChanged, IDisposable
    {
        private readonly IDownloadService _downloadService;
        private readonly IDuplicateDetectionService _duplicateDetectionService;
        private readonly IProgressTrackingService _progressTrackingService;
        private readonly ILoggingService _loggingService;
        private readonly IUIService _uiService;
        private readonly object _queueLock = new();

        private string _downloadPath = string.Empty;
        private int _maxConcurrentDownloads = 4;
        private string _statusSummary = string.Empty;

        private long _totalDownloadedBytes;
        private int _completedDownloads;
        private bool _isDownloadStarted;
        private bool _disposed;

        private readonly ConcurrentQueue<VideoItem> _downloadQueue = new();
        private int _activeDownloadCount;
        private readonly SemaphoreSlim _queueMonitorSignal;
        private readonly List<Task> _activeDownloadTasks = new();
        private CancellationTokenSource _globalCts;
        private readonly CancellationTokenSource _queueCts;

        public ObservableCollection<VideoItem> VideoItems { get; } = new();

        public int MaxConcurrentDownloads
        {
            get => _maxConcurrentDownloads;
            set
            {
                if (SetField(ref _maxConcurrentDownloads, value))
                {
                    _downloadService.SetMaxConcurrentDownloads(value);
                }
            }
        }

        public string StatusSummary
        {
            get => _statusSummary;
            private set => SetField(ref _statusSummary, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public DownloadManager(
            IDownloadService downloadService,
            IDuplicateDetectionService duplicateDetectionService,
            IProgressTrackingService progressTrackingService,
            ILoggingService loggingService,
            IUIService? uiService = null)
        {
            _downloadService = downloadService;
            _duplicateDetectionService = duplicateDetectionService;
            _progressTrackingService = progressTrackingService;
            _loggingService = loggingService;
            _uiService = uiService ?? new UIService();
            _queueMonitorSignal = new SemaphoreSlim(0, int.MaxValue);
            _globalCts = new CancellationTokenSource();
            _queueCts = new CancellationTokenSource();
            
            VideoItems.CollectionChanged += (s, e) => UpdateStatusSummary();
        }

        public void Initialize(string downloadPath)
        {
            lock (_queueLock)
            {
                _downloadPath = downloadPath;
                _duplicateDetectionService.SetDownloadPath(downloadPath);
                StartQueueMonitor();
                _loggingService.Info("DownloadManager初始化完成", "DownloadManager");
            }
        }

        public void StartDownload(VideoItem item)
        {
            _isDownloadStarted = true;
            AddToQueue(item);
        }

        private void AddToQueue(VideoItem item)
        {
            _downloadQueue.Enqueue(item);
            _queueMonitorSignal.Release();
        }

        private void StartQueueMonitor()
        {
            Task.Run(MonitorDownloadQueue);
        }

        private async Task MonitorDownloadQueue()
        {
            while (!_queueCts.IsCancellationRequested)
            {
                try
                {
                    await _queueMonitorSignal.WaitAsync(500, _queueCts.Token);

                    while (_activeDownloadCount < MaxConcurrentDownloads)
                    {
                        if (!_downloadQueue.TryDequeue(out var item)) break;

                        if (item != null &&
                            item.StatusFlag != DownloadStatus.Downloading &&
                            item.StatusFlag != DownloadStatus.Completed)
                        {
                            StartDownloadTask(item);
                        }
                    }

                    CheckDownloadCompletion();
                }
                catch (OperationCanceledException) when (_queueCts.IsCancellationRequested)
                {
                    break;
                }
                catch
                {
                    await Task.Delay(500, _queueCts.Token);
                }
            }
        }

        private void StartDownloadTask(VideoItem item)
        {
            try
            {
                UpdateItemStatus(item, "下载中...", "#FFFFA500", DownloadStatus.Downloading);

                item.DownloadTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_globalCts.Token);
                var token = item.DownloadTokenSource.Token;

                Interlocked.Increment(ref _activeDownloadCount);

                var task = Task.Run(async () =>
                {
                    try
                    {
                        var result = await DownloadVideoAsync(item, token);
                        return result;
                    }
                    finally
                    {
                        Interlocked.Decrement(ref _activeDownloadCount);
                        item.DownloadTokenSource?.Dispose();
                        item.DownloadTokenSource = null;
                        _queueMonitorSignal.Release();

                        lock (_queueLock)
                        {
                            _activeDownloadTasks.RemoveAll(t => t.Id == Task.CurrentId);
                        }
                    }
                }, token);

                lock (_queueLock)
                {
                    _activeDownloadTasks.Add(task);
                }
            }
            catch (Exception ex)
            {
                _loggingService.Error($"启动下载失败: {ex.Message}", "DownloadManager");
            }
        }

        private async Task<bool> DownloadVideoAsync(VideoItem item, CancellationToken token)
        {
            if (item == null) return false;

            var safeName = SanitizeFileName(item.Name);
            var outputFile = Path.Combine(_downloadPath, $"{safeName}.mp4");

            if (string.IsNullOrEmpty(safeName))
            {
                UpdateItemStatus(item, "下载失败: 无效文件名", "#FFFF0000", DownloadStatus.Failed);
                return false;
            }

            if (outputFile.Length > 260)
            {
                UpdateItemStatus(item, "下载失败: 路径过长", "#FFFF0000", DownloadStatus.Failed);
                return false;
            }

            int retryCount = 0;
            const int maxRetries = 3;

            using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(300));
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(token, timeoutCts.Token);

            while (retryCount <= maxRetries)
            {
                try
                {
                    if (string.IsNullOrEmpty(item.Url))
                    {
                        UpdateItemStatus(item, "下载失败: 无效URL", "#FFFF0000", DownloadStatus.Failed);
                        return false;
                    }

                    var fileState = await CheckExistingFileStateAsync(item, outputFile);
                    if (fileState == FileDownloadState.Completed)
                    {
                        return true;
                    }

                    item.LastReportedBytes = item.DownloadedBytes;

                    void ProgressCallback(long downloaded, long total, double speed, TimeSpan remaining)
                    {
                        UpdateDownloadProgress(item, downloaded, total, speed, remaining);
                    }

                    var result = await _downloadService.DownloadWithConcurrencyControl(
                        item.Url,
                        outputFile,
                        ProgressCallback,
                        linkedCts.Token,
                        () => _loggingService.Info($"开始下载: {item.Name}", "DownloadManager"));

                    if (result)
                    {
                        await Task.Delay(500, linkedCts.Token);
                        var fileInfo = new FileInfo(outputFile);

                        UpdateItemStatus(item, "下载完成", "#FF008000", DownloadStatus.Completed);
                        item.Progress = 100;
                        item.DownloadedBytes = item.TotalBytes;

                        Interlocked.Increment(ref _completedDownloads);
                        _duplicateDetectionService.AddToCache(item);
                        _loggingService.Info($"下载完成：{item.Name} ({FormatUtils.FormatSize(fileInfo.Length)})", "DownloadManager");
                        return true;
                    }
                    else
                    {
                        throw new Exception("下载返回失败结果");
                    }
                }
                catch (OperationCanceledException)
                {
                    if (File.Exists(outputFile) && new FileInfo(outputFile).Length > 0)
                    {
                        UpdateItemStatus(item, "已暂停", "#FFFFA500", DownloadStatus.Partial);
                    }
                    else
                    {
                        UpdateItemStatus(item, "已取消", "#FFFFA500", DownloadStatus.Pending);
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    retryCount++;
                    if (retryCount > maxRetries)
                    {
                        _loggingService.Error($"下载失败: {ex.Message}", "DownloadManager");
                        UpdateItemStatus(item, "下载失败", "#FFFF0000", DownloadStatus.Failed);
                        return false;
                    }

                    var delay = (int)Math.Pow(2, retryCount) * 1000;
                    _loggingService.Warning($"下载失败，将在 {delay}ms 后重试 ({retryCount}/{maxRetries}): {ex.Message}", "DownloadManager");
                    await Task.Delay(delay, linkedCts.Token);
                }
            }

            return false;
        }

        private enum FileDownloadState
        {
            None,
            Partial,
            Completed
        }

        private async Task<FileDownloadState> CheckExistingFileStateAsync(VideoItem item, string outputFile)
        {
            if (!File.Exists(outputFile))
            {
                return FileDownloadState.None;
            }

            var fileInfo = new FileInfo(outputFile);
            var fileSize = fileInfo.Length;

            if (item.FileSize > 0 && fileSize == item.FileSize)
            {
                UpdateItemStatus(item, "下载完成", "#FF008000", DownloadStatus.Completed);
                item.Progress = 100;
                Interlocked.Increment(ref _completedDownloads);
                return FileDownloadState.Completed;
            }

            if (fileSize > 0)
            {
                UpdateItemStatus(item, "继续下载", "#FF0000FF", DownloadStatus.Partial);
                item.DownloadedBytes = fileSize;
                item.Progress = item.FileSize > 0 ? (int)(fileSize * 100 / item.FileSize) : 0;
                return FileDownloadState.Partial;
            }

            return FileDownloadState.None;
        }

        private void UpdateItemStatus(VideoItem item, string status, string colorHex, DownloadStatus flag)
        {
            _uiService.Invoke(() =>
            {
                item.Status = status;
                item.StatusColorHex = colorHex;
                item.StatusFlag = flag;
            });
        }

        private readonly System.Diagnostics.Stopwatch _uiUpdateStopwatch = System.Diagnostics.Stopwatch.StartNew();
        private int _uiUpdateCounter;

        private void UpdateDownloadProgress(VideoItem item, long downloaded, long total, double speed, TimeSpan remaining)
        {
            const int throttleMs = 200;
            if (_uiUpdateStopwatch.ElapsedMilliseconds < throttleMs) return;
            _uiUpdateStopwatch.Restart();

            _uiService.Invoke(() =>
            {
                if (total > 0 && downloaded >= total)
                {
                    item.Progress = 100;
                    item.DownloadedBytes = total;
                    item.TotalBytes = total;
                    item.CurrentSpeedBytesPerSec = 0;
                    item.RemainingTime = TimeSpan.Zero;
                }
                else
                {
                    var delta = downloaded - item.LastReportedBytes;
                    item.LastReportedBytes = downloaded;
                    Interlocked.Add(ref _totalDownloadedBytes, delta);
                    item.DownloadedBytes = downloaded;
                    item.TotalBytes = total;
                    if (total > 0) item.Progress = (int)(downloaded * 100 / total);
                    item.CurrentSpeedBytesPerSec = speed;
                    item.RemainingTime = remaining;
                }

                // 每3次更新才更新一次状态摘要，减少UI更新频率
                if (Interlocked.Increment(ref _uiUpdateCounter) % 3 == 0)
                {
                    UpdateStatusSummary();
                }
            });
        }

        private void UpdateStatusSummary()
        {
            var progressInfo = _progressTrackingService.CalculateOverallProgress(VideoItems);

            if (!progressInfo.AnyActive && _isDownloadStarted)
            {
                StatusSummary = "等待队列中的任务开始下载...";
            }
            else if (!progressInfo.AnyActive)
            {
                StatusSummary = VideoItems.Count > 0 ? "当前没有视频正在下载" : "当前没有视频正在下载，统计信息为空";
            }
            else
            {
                StatusSummary = $"当前下载速度: {_progressTrackingService.FormatSpeed(progressInfo.CurrentSpeedBytesPerSecond)} | " +
                    $"已下载: {progressInfo.CompletedCount}/{VideoItems.Count} | " +
                    $"下载完成: {progressInfo.OverallPercentage:F1}% | " +
                    $"活动下载数: {progressInfo.ActiveDownloadCount}/{MaxConcurrentDownloads}";
            }
        }

        private void CheckDownloadCompletion()
        {
            if (_downloadQueue.IsEmpty && _activeDownloadCount == 0 && _isDownloadStarted)
            {
                _isDownloadStarted = false;
                var completedCount = Interlocked.CompareExchange(ref _completedDownloads, 0, 0);
                _loggingService.Info($"全部视频下载完成，成功：{completedCount} 个", "DownloadManager");
            }
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return name;

            var invalidChars = Path.GetInvalidFileNameChars();
            var builder = new System.Text.StringBuilder();

            foreach (var c in name)
            {
                builder.Append(Array.IndexOf(invalidChars, c) >= 0 ? '_' : c);
            }

            return builder.ToString();
        }

        public void ClearAll()
        {
            CancelAllDownloads();
            _downloadQueue.Clear();
            _activeDownloadCount = 0;
            _isDownloadStarted = false;
            _totalDownloadedBytes = 0;
            _completedDownloads = 0;

            lock (_queueLock)
            {
                _activeDownloadTasks.Clear();
            }

            _duplicateDetectionService.ClearCache();
        }

        public void CancelAllDownloads()
        {
            var newCts = new CancellationTokenSource();

            lock (_queueLock)
            {
                var oldCts = _globalCts;
                _globalCts = newCts;

                try { oldCts.Cancel(); } catch { }
                try { oldCts.Dispose(); } catch { }
            }

            foreach (var item in VideoItems)
            {
                if (item.StatusFlag == DownloadStatus.Downloading)
                {
                    item.DownloadTokenSource?.Cancel();
                    item.Status = "已取消";
                    item.StatusColorHex = "#FFFFA500";
                    item.StatusFlag = DownloadStatus.Pending;
                }
            }
        }

        public void UpdatePath(string newPath)
        {
            _downloadPath = newPath;
            _duplicateDetectionService.SetDownloadPath(newPath);
        }

        public void AddVideoItem(VideoItem item)
        {
            if (!_duplicateDetectionService.IsDuplicate(item.Name))
            {
                _duplicateDetectionService.AddToCache(item.Name);
                VideoItems.Add(item);
                UpdateVideoIndexes();
                UpdateStatusSummary();
            }
        }

        private void UpdateVideoIndexes()
        {
            for (int i = 0; i < VideoItems.Count; i++)
            {
                VideoItems[i].DisplayIndex = i + 1;
            }
        }

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            _queueCts.Cancel();
            _queueCts.Dispose();

            _globalCts?.Dispose();
            _queueMonitorSignal?.Dispose();

            _loggingService?.Dispose();
        }
    }
}