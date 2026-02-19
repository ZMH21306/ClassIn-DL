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
using Classin视频解析下载工具.Services.CoreServices;
using Classin视频解析下载工具.Services.DownloadServices;
using Classin视频解析下载工具.Services.UIServices;
using Classin视频解析下载工具.Shared.Helpers;

namespace Classin视频解析下载工具.Services.DownloadServices
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
        private string _statusSummary = string.Empty;

        private long _totalDownloadedBytes;
        private int _completedDownloads;
        private bool _isDownloadStarted;
        private bool _disposed;

        private readonly Queue<VideoItem> _downloadQueue = new();
        private bool _isDownloading;
        private Task? _currentDownloadTask;
        private CancellationTokenSource _globalCts;

        public ObservableCollection<VideoItem> VideoItems { get; } = new();

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
            _globalCts = new CancellationTokenSource();

            VideoItems.CollectionChanged += (s, e) => UpdateStatusSummary();
        }

        public void Initialize(string downloadPath)
        {
            lock (_queueLock)
            {
                _downloadPath = downloadPath;
                _duplicateDetectionService.SetDownloadPath(downloadPath);
                _loggingService.Info("DownloadManager初始化完成", "DownloadManager");
            }
        }

        public void StartDownload(VideoItem item)
        {
            _loggingService.Info($"开始下载视频项: {item.Name}", "DownloadManager");
            _isDownloadStarted = true;
            AddToQueue(item);
        }

        private void AddToQueue(VideoItem item)
        {
            _loggingService.Debug($"将视频项添加到下载队列: {item.Name}", "DownloadManager.Queue");

            // 设置视频项状态为等待下载
            UpdateItemStatus(item, "等待下载", "#FF1E90FF", DownloadStatus.Pending);

            lock (_queueLock)
            {
                _downloadQueue.Enqueue(item);
            }

            ProcessQueue();
        }

        private void ProcessQueue()
        {
            if (_isDownloading) return;

            lock (_queueLock)
            {
                if (_downloadQueue.Count == 0 || _isDownloading) return;

                if (_downloadQueue.TryDequeue(out var item))
                {
                    StartDownloadTask(item);
                }
            }
        }

        private void StartDownloadTask(VideoItem item)
        {
            try
            {
                _loggingService.Info($"开始执行下载任务: {item.Name}", "DownloadManager.Task");

                UpdateItemStatus(item, "下载中...", "#FFFFA500", DownloadStatus.Downloading);

                item.DownloadTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_globalCts.Token);
                var token = item.DownloadTokenSource.Token;

                _isDownloading = true;
                _loggingService.Debug($"开始下载任务", "DownloadManager.Task");

                _currentDownloadTask = Task.Run(async () =>
                {
                    try
                    {
                        _loggingService.Debug($"开始异步下载: {item.Name}", "DownloadManager.Download");

                        var result = await DownloadVideoAsync(item, token);
                        return result;
                    }
                    finally
                    {
                        _isDownloading = false;
                        item.DownloadTokenSource?.Dispose();
                        item.DownloadTokenSource = null;

                        _loggingService.Debug($"下载任务完成", "DownloadManager.Task");

                        ProcessQueue();
                    }
                }, token);

                _loggingService.Debug($"下载任务已启动", "DownloadManager.Task");
            }
            catch (Exception ex)
            {
                _loggingService.Error($"启动下载任务时发生异常: {ex.Message}", "DownloadManager.Task");
                UpdateItemStatus(item, "启动失败", "#FFFF0000", DownloadStatus.Failed);
                _isDownloading = false;
                ProcessQueue();
            }
        }

        private async Task<bool> DownloadVideoAsync(VideoItem item, CancellationToken token)
        {
            if (item == null) return false;

            _loggingService.Debug($"开始准备下载: {item.Name}", "DownloadManager.Download");

            var safeName = SanitizeFileName(item.Name);
            var outputFile = Path.Combine(_downloadPath, $"{safeName}.mp4");

            if (string.IsNullOrEmpty(safeName))
            {
                _loggingService.Error($"下载失败：无效文件名 {item.Name}", "DownloadManager.Download");
                UpdateItemStatus(item, "下载失败: 无效文件名", "#FFFF0000", DownloadStatus.Failed);
                return false;
            }

            if (outputFile.Length > 260)
            {
                _loggingService.Error($"下载失败：路径过长 {outputFile}", "DownloadManager.Download");
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
                        _loggingService.Error($"下载失败：无效URL {item.Name}", "DownloadManager.Download");
                        UpdateItemStatus(item, "下载失败: 无效URL", "#FFFF0000", DownloadStatus.Failed);
                        return false;
                    }

                    _loggingService.Debug($"检查现有文件状态: {outputFile}", "DownloadManager.Download");
                    var fileState = await CheckExistingFileStateAsync(item, outputFile);
                    if (fileState == FileDownloadState.Completed)
                    {
                        _loggingService.Info($"文件已存在且完整，跳过下载: {item.Name}", "DownloadManager.Download");
                        return true;
                    }

                    item.LastReportedBytes = item.DownloadedBytes;

                    void ProgressCallback(long downloaded, long total, double speed, TimeSpan remaining)
                    {
                        UpdateDownloadProgress(item, downloaded, total, speed, remaining);
                    }

                    _loggingService.Info($"开始下载文件: {item.Name} -> {outputFile}", "DownloadManager.Download");
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

                        _uiService.Invoke(() =>
                        {
                            item.Status = "下载完成";
                            item.StatusColorHex = "#FF008000";
                            item.StatusFlag = DownloadStatus.Completed;
                            item.Progress = 100;
                            item.DownloadedBytes = item.TotalBytes;
                            item.CurrentSpeedBytesPerSec = 0;
                            item.RemainingTime = TimeSpan.Zero;
                        });

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
                    _loggingService.Info($"下载被用户取消: {item.Name}", "DownloadManager.Download");
                    if (File.Exists(outputFile) && new FileInfo(outputFile).Length > 0)
                    {
                        UpdateItemStatus(item, "已暂停", "#FF808080", DownloadStatus.Partial);
                    }
                    else
                    {
                        UpdateItemStatus(item, "已取消", "#FF808080", DownloadStatus.Pending);
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    retryCount++;
                    _loggingService.Error($"下载过程中发生异常: {ex.Message}，正在进行第 {retryCount} 次重试: {item.Name}", "DownloadManager.Download");

                    if (retryCount > maxRetries)
                    {
                        _loggingService.Error($"下载失败，已达到最大重试次数: {item.Name}", "DownloadManager.Download");
                        UpdateItemStatus(item, "下载失败", "#FFFF0000", DownloadStatus.Failed);
                        return false;
                    }

                    var delay = (int)Math.Pow(2, retryCount) * 1000;
                    _loggingService.Warning($"下载失败，将在 {delay}ms 后重试 ({retryCount}/{maxRetries}): {ex.Message}", "DownloadManager.Download");
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
            await Task.CompletedTask;

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
                UpdateItemStatus(item, "继续下载", "#FF808080", DownloadStatus.Partial);
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

        private void UpdateDownloadProgress(VideoItem item, long downloaded, long total, double speed, TimeSpan remaining)
        {
            var now = System.Diagnostics.Stopwatch.GetTimestamp();
            var elapsedMs = (now - item.LastProgressUpdateTimestamp) * 1000.0 / System.Diagnostics.Stopwatch.Frequency;

            // 增加时间间隔，减少更新频率
            if (elapsedMs < 500 && downloaded < total) return;

            item.LastProgressUpdateTimestamp = now;

            _uiService.Invoke(() =>
            {
                if (total > 0 && downloaded >= total)
                {
                    // 一次性更新所有属性，减少UI通知次数
                    item.UpdateProgress(total, total, 100, 0, TimeSpan.Zero);
                }
                else
                {
                    var delta = downloaded - item.LastReportedBytes;
                    item.LastReportedBytes = downloaded;
                    Interlocked.Add(ref _totalDownloadedBytes, delta);

                    // 使用统一的进度更新方法，减少UI通知次数
                    int progress = total > 0 ? (int)(downloaded * 100 / total) : 0;
                    item.UpdateProgress(downloaded, total, progress, speed, remaining);
                }

                // 减少状态摘要更新频率
                if (elapsedMs >= 1000) // 每秒更新一次状态摘要
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
                    $"下载完成: {progressInfo.OverallPercentage:F1}%";
            }
        }

        private void CheckDownloadCompletion()
        {
            if (_downloadQueue.Count == 0 && !_isDownloading && _isDownloadStarted)
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
            _loggingService.Info("开始清空所有下载任务和数据", "DownloadManager");
            CancelAllDownloads();
            _downloadQueue.Clear();
            _isDownloading = false;
            _isDownloadStarted = false;
            _totalDownloadedBytes = 0;
            _completedDownloads = 0;

            _duplicateDetectionService.ClearCache();
            VideoItems.Clear();

            _loggingService.Info("所有下载任务和数据已清空", "DownloadManager");
        }

        public void CancelAllDownloads()
        {
            _loggingService.Info("开始取消所有下载任务", "DownloadManager");

            var newCts = new CancellationTokenSource();

            lock (_queueLock)
            {
                var oldCts = _globalCts;
                _globalCts = newCts;

                try { oldCts.Cancel(); } catch { }
                try { oldCts.Dispose(); } catch { }
            }

            int cancelledCount = 0;
            foreach (var item in VideoItems)
            {
                if (item.StatusFlag == DownloadStatus.Downloading)
                {
                    item.DownloadTokenSource?.Cancel();
                    item.Status = "已取消";
                    item.StatusColorHex = "#FFFFA500";
                    item.StatusFlag = DownloadStatus.Pending;
                    cancelledCount++;
                }
            }

            _loggingService.Info($"已取消 {cancelledCount} 个下载任务", "DownloadManager");
        }

        public void UpdatePath(string newPath)
        {
            _downloadPath = newPath;
            _duplicateDetectionService.SetDownloadPath(newPath);
        }

        // 批量添加视频项，减少UI更新次数
        public void AddVideoItems(IEnumerable<VideoItem> items)
        {
            // 使用预分配容量的List，减少内存重分配
            var newItems = new List<VideoItem>();

            foreach (var item in items)
            {
                if (!_duplicateDetectionService.IsDuplicate(item.Name))
                {
                    _duplicateDetectionService.AddToCache(item.Name);
                    newItems.Add(item);
                }
            }

            if (newItems.Count > 0)
            {
                _uiService.Invoke(() =>
                {
                    // 批量添加，减少UI通知次数
                    foreach (var item in newItems)
                    {
                        VideoItems.Add(item);
                    }

                    // 一次性更新所有索引
                    UpdateVideoIndexes();

                    // 更新状态摘要
                    UpdateStatusSummary();
                });
            }
        }

        public void AddVideoItem(VideoItem item)
        {
            if (!_duplicateDetectionService.IsDuplicate(item.Name))
            {
                _duplicateDetectionService.AddToCache(item.Name);
                _uiService.Invoke(() =>
                {
                    VideoItems.Add(item);
                    UpdateVideoIndexes();
                    UpdateStatusSummary();
                });
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

            _globalCts?.Cancel();
            _globalCts?.Dispose();

            _loggingService?.Dispose();
        }
    }
}