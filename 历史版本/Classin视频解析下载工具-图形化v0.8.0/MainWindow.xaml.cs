using Microsoft.Win32;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace VideoDownloader
{
    /// <summary>
    /// 索引格式转换器，将整数索引转换为两位数格式字符串
    /// </summary>
    public class IndexFormatConverter : IValueConverter
    {
        /// <summary>
        /// 将整数索引转换为两位数格式字符串
        /// </summary>
        /// <param name="value">要转换的索引值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="parameter">转换参数</param>
        /// <param name="culture">文化信息</param>
        /// <returns>两位数格式的索引字符串</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is int index ? index.ToString("D2") : "00";
        }

        /// <summary>
        /// 反向转换（未实现）
        /// </summary>
        /// <param name="value">要转换的值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="parameter">转换参数</param>
        /// <param name="culture">文化信息</param>
        /// <returns>转换结果</returns>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// 高度转换器，根据输入高度计算合适的项高度
    /// </summary>
    public class HeightConverter : IValueConverter
    {
        /// <summary>
        /// 根据输入高度计算合适的项高度，限制在30-80之间
        /// </summary>
        /// <param name="value">输入高度值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="parameter">转换参数</param>
        /// <param name="culture">文化信息</param>
        /// <returns>计算后的项高度</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double height)
            {
                // 计算项高度：(输入高度 - 5) / 10
                double itemHeight = (height - 5) / 10.0;
                // 限制高度在30-80之间
                return Math.Clamp(itemHeight, 30, 80);
            }
            // 默认高度为40
            return 40.0;
        }

        /// <summary>
        /// 反向转换（未实现）
        /// </summary>
        /// <param name="value">要转换的值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="parameter">转换参数</param>
        /// <param name="culture">文化信息</param>
        /// <returns>转换结果</returns>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// 颜色变暗转换器，根据指定的变暗因子将颜色变暗
    /// </summary>
    public class DarkenColorConverter : IValueConverter
    {
        /// <summary>
        /// 变暗因子，默认值为0.3（变暗30%）
        /// </summary>
        public double DarkenFactor { get; set; } = 0.3;

        /// <summary>
        /// 将颜色根据变暗因子进行变暗处理
        /// </summary>
        /// <param name="value">要变暗的SolidColorBrush</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="parameter">转换参数</param>
        /// <param name="culture">文化信息</param>
        /// <returns>变暗后的SolidColorBrush</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is SolidColorBrush brush)
            {
                Color color = brush.Color;
                // 计算变暗后的RGB值：原值 * (1 - 变暗因子)
                byte r = (byte)(color.R * (1 - DarkenFactor));
                byte g = (byte)(color.G * (1 - DarkenFactor));
                byte b = (byte)(color.B * (1 - DarkenFactor));
                return new SolidColorBrush(Color.FromRgb(r, g, b));
            }
            // 如果输入不是SolidColorBrush，直接返回原值
            return value;
        }

        /// <summary>
        /// 反向转换（未实现）
        /// </summary>
        /// <param name="value">要转换的值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="parameter">转换参数</param>
        /// <param name="culture">文化信息</param>
        /// <returns>转换结果</returns>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// 缩放转换器，根据当前宽度与基础宽度的比例计算缩放尺寸
    /// </summary>
    public class ScalingConverter : IValueConverter
    {
        /// <summary>
        /// 基础宽度，默认值为1440
        /// </summary>
        public double BaseWidth { get; set; } = 1440;

        /// <summary>
        /// 根据当前宽度与基础宽度的比例计算缩放尺寸，缩放比例限制在0.7-1.3之间
        /// </summary>
        /// <param name="value">当前宽度值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="parameter">基础尺寸参数（字符串形式）</param>
        /// <param name="culture">文化信息</param>
        /// <returns>缩放后的尺寸</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double width && parameter is string paramString && double.TryParse(paramString, out double baseSize))
            {
                // 计算缩放比例：当前宽度 / 基础宽度，限制在0.7-1.3之间
                double scale = Math.Clamp(width / BaseWidth, 0.7, 1.3);
                // 返回缩放后的尺寸
                return baseSize * scale;
            }
            // 如果转换条件不满足，返回原始参数
            return parameter;
        }

        /// <summary>
        /// 反向转换（未实现）
        /// </summary>
        /// <param name="value">要转换的值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="parameter">转换参数</param>
        /// <param name="culture">文化信息</param>
        /// <returns>转换结果</returns>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// 路径最大宽度转换器，根据网格宽度、当前宽度和按钮宽度计算路径显示的最大宽度
    /// </summary>
    public class PathMaxWidthConverter : IMultiValueConverter
    {
        /// <summary>
        /// 根据网格宽度、当前宽度和按钮宽度计算路径显示的最大宽度
        /// </summary>
        /// <param name="values">值数组，包含：[网格宽度, 当前宽度, 按钮宽度]</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="parameter">转换参数</param>
        /// <param name="culture">文化信息</param>
        /// <returns>路径显示的最大宽度</returns>
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            // 检查值数组是否包含至少3个有效元素
            if (values.Length < 3 ||
                !(values[0] is double gridWidth) ||
                !(values[1] is double currentWidth) ||
                !(values[2] is double buttonWidth))
            {
                // 默认最大宽度为300
                return 300;
            }

            // 计算可用宽度：
            // 网格宽度 - 网格宽度的15% - 按钮宽度 - 40（额外边距）
            double availableWidth = gridWidth -
                                   (gridWidth * 0.15) -
                                   buttonWidth -
                                   40;

            // 返回可用宽度，最小为100
            return Math.Max(100, availableWidth);
        }

        /// <summary>
        /// 反向转换（未实现）
        /// </summary>
        /// <param name="value">要转换的值</param>
        /// <param name="targetTypes">目标类型数组</param>
        /// <param name="parameter">转换参数</param>
        /// <param name="culture">文化信息</param>
        /// <returns>转换结果数组</returns>
        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public partial class MainWindow : Window
    {
        private const int DEFAULT_MAX_CONCURRENT_DOWNLOADS = 5;
        private const int DEFAULT_MAX_DOWNLOAD_THREADS = 32;

        private string downloadPath = string.Empty;
        private ObservableCollection<VideoItem> videoItems = new ObservableCollection<VideoItem>();
        private CancellationTokenSource? _clipboardCts;
        private bool _isClosing = false;
        private bool _isDownloadStarted = false;

        private double[] columnRatios = { 0.05, 0.3, 0.4, 0.25 };
        private bool userAdjustedColumns = false;
        private double lastTotalWidth = 0;

        private GridLength originalPanelWidth;
        private GridLength originalResultWidth;
        private GridLength originalMainContentHeight;
        private GridLength originalLogHeight;
        private double[] originalColumnRatios = { 0.05, 0.3, 0.4, 0.25 };

        private bool _logScrollToEnd = true;
        private bool _isInitialized = false;

        private long _totalBytesToDownload = 0;
        private long _totalDownloadedBytes = 0;
        private int _completedDownloads = 0;
        private DispatcherTimer _statusTimer;

        private CancellationTokenSource _globalCts = new CancellationTokenSource();
        private ConcurrentBag<Task> _activeDownloadTasks = new ConcurrentBag<Task>();

        private DateTime _lastUiUpdate = DateTime.MinValue;
        private const int UI_UPDATE_THROTTLE_MS = 100;

        private int _activeDownloadCount = 0;
        private ConcurrentQueue<VideoItem> _downloadQueue = new ConcurrentQueue<VideoItem>();
        private SemaphoreSlim _slotMonitor = new SemaphoreSlim(1, 1);
        private CancellationTokenSource _queueCts = new CancellationTokenSource();

        private SemaphoreSlim _queueMonitorSignal = new SemaphoreSlim(0, int.MaxValue);

        private int _lastReportedConcurrentDownloads = DEFAULT_MAX_CONCURRENT_DOWNLOADS;
        private int _lastReportedDownloadThreads = DEFAULT_MAX_DOWNLOAD_THREADS;

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern int DeleteMenu(IntPtr hMenu, int nPosition, int wFlags);
        [DllImport("user32.dll")]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        private const int SC_CLOSE = 0xF060;
        private const int MF_BYCOMMAND = 0x00000000;

        /// <summary>
        /// 视频项类，用于存储和管理单个视频的下载信息
        /// </summary>
        public class VideoItem : INotifyPropertyChanged
        {
            /// <summary>
            /// 下载状态枚举
            /// </summary>
            public enum DownloadStatus
            {
                /// <summary>等待下载</summary>
                Pending,
                /// <summary>正在下载</summary>
                Downloading,
                /// <summary>下载完成</summary>
                Completed,
                /// <summary>下载失败</summary>
                Failed,
                /// <summary>下载暂停</summary>
                Partial
            }

            private DownloadStatus _statusFlag = DownloadStatus.Pending;
            /// <summary>
            /// 下载状态标志
            /// </summary>
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

            /// <summary>
            /// 是否为活动下载（正在下载状态）
            /// </summary>
            public bool IsActiveDownload => StatusFlag == DownloadStatus.Downloading;

            private int _displayIndex;
            /// <summary>
            /// 显示索引
            /// </summary>
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

            /// <summary>
            /// 视频名称
            /// </summary>
            public string Name { get; set; } = string.Empty;
            /// <summary>
            /// 视频下载链接
            /// </summary>
            public string Url { get; set; } = string.Empty;

            private long _fileSize;
            /// <summary>
            /// 文件大小
            /// </summary>
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
            /// <summary>
            /// 状态描述
            /// </summary>
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
            /// <summary>
            /// 下载进度（0-100）
            /// </summary>
            public int Progress
            {
                get => _progress;
                set
                {
                    if (_progress != value)
                    {
                        _progress = value;
                        OnPropertyChanged(nameof(Progress));
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            private long _downloadedBytes;
            /// <summary>
            /// 已下载字节数
            /// </summary>
            public long DownloadedBytes
            {
                get => _downloadedBytes;
                set
                {
                    if (_downloadedBytes != value)
                    {
                        _downloadedBytes = value;
                        OnPropertyChanged(nameof(DownloadedBytes));
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            private long _totalBytes;
            /// <summary>
            /// 总字节数
            /// </summary>
            public long TotalBytes
            {
                get => _totalBytes;
                set
                {
                    if (_totalBytes != value)
                    {
                        _totalBytes = value;
                        OnPropertyChanged(nameof(TotalBytes));
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            private double _currentSpeedBytesPerSec;
            /// <summary>
            /// 当前下载速度（字节/秒）
            /// </summary>
            public double CurrentSpeedBytesPerSec
            {
                get => _currentSpeedBytesPerSec;
                set
                {
                    if (_currentSpeedBytesPerSec != value)
                    {
                        _currentSpeedBytesPerSec = value;
                        OnPropertyChanged(nameof(CurrentSpeedBytesPerSec));
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            private TimeSpan _remainingTime = TimeSpan.MaxValue;
            /// <summary>
            /// 剩余下载时间
            /// </summary>
            public TimeSpan RemainingTime
            {
                get => _remainingTime;
                set
                {
                    if (_remainingTime != value)
                    {
                        _remainingTime = value;
                        OnPropertyChanged(nameof(RemainingTime));
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            private string _phase = "等待";
            /// <summary>
            /// 当前下载阶段
            /// </summary>
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
            /// <summary>
            /// 上次报告的字节数，用于计算下载速度
            /// </summary>
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

            /// <summary>
            /// 下载取消令牌源
            /// </summary>
            public CancellationTokenSource? DownloadTokenSource { get; set; }

            /// <summary>
            /// 显示状态，根据下载状态生成格式化的状态描述
            /// </summary>
            public string DisplayStatus
            {
                get
                {
                    switch (StatusFlag)
                    {
                        case DownloadStatus.Downloading:
                            return $"下载中 ({Progress}%) - {FormatSize(DownloadedBytes)}/{FormatSize(TotalBytes)} @ {FormatSpeed(CurrentSpeedBytesPerSec)} - 剩余: {FormatTime(RemainingTime)}";
                        case DownloadStatus.Partial:
                            return $"已暂停 ({Progress}%) - {FormatSize(DownloadedBytes)}/{FormatSize(TotalBytes)}";
                        case DownloadStatus.Completed:
                            return "下载完成";
                        case DownloadStatus.Failed:
                            return "下载失败";
                        default:
                            return Status;
                    }
                }
            }

            /// <summary>
            /// 格式化字节大小为可读字符串
            /// </summary>
            /// <param name="bytes">字节大小</param>
            /// <returns>可读的大小字符串（B/KB/MB/GB）</returns>
            public static string FormatSize(long bytes)
            {
                if (bytes >= 1 << 30) return $"{bytes / (1 << 30):F2} GB"; // 大于等于1GB
                if (bytes >= 1 << 20) return $"{bytes / (1 << 20):F2} MB"; // 大于等于1MB
                if (bytes >= 1 << 10) return $"{bytes / (1 << 10):F2} KB"; // 大于等于1KB
                return $"{bytes} B"; // 小于1KB
            }

            /// <summary>
            /// 格式化下载速度为可读字符串
            /// </summary>
            /// <param name="bytesPerSecond">字节/秒</param>
            /// <returns>可读的速度字符串（B/s/KB/s/MB/s/GB/s）</returns>
            public static string FormatSpeed(double bytesPerSecond)
            {
                if (bytesPerSecond >= 1 << 30) return $"{bytesPerSecond / (1 << 30):F2} GB/s"; // 大于等于1GB/s
                if (bytesPerSecond >= 1 << 20) return $"{bytesPerSecond / (1 << 20):F2} MB/s"; // 大于等于1MB/s
                if (bytesPerSecond >= 1 << 10) return $"{bytesPerSecond / (1 << 10):F2} KB/s"; // 大于等于1KB/s
                return $"{bytesPerSecond:F2} B/s"; // 小于1KB/s
            }

            /// <summary>
            /// 格式化时间间隔为可读字符串
            /// </summary>
            /// <param name="time">时间间隔</param>
            /// <returns>可读的时间字符串（HH:MM:SS）</returns>
            public static string FormatTime(TimeSpan time)
            {
                if (time == TimeSpan.MaxValue) return "未知"; // 未知时间
                return $"{time.Hours:D2}:{time.Minutes:D2}:{time.Seconds:D2}"; // 格式化为HH:MM:SS
            }

            /// <summary>
            /// 状态颜色
            /// </summary>
            public Brush StatusColor { get; set; } = Brushes.Gray;

            /// <summary>
            /// 属性更改事件
            /// </summary>
            public event PropertyChangedEventHandler? PropertyChanged;

            /// <summary>
            /// 触发属性更改事件
            /// </summary>
            /// <param name="propertyName">属性名称</param>
            protected virtual void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public MainWindow()
        {
            InitializeComponent();

            downloadPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "下载目录");
            Directory.CreateDirectory(downloadPath);

            this.Closed += MainWindow_Closed;

            originalPanelWidth = PanelColumn.Width;
            originalResultWidth = ResultColumn.Width;
            originalMainContentHeight = MainContentRow.Height;
            originalLogHeight = LogRow.Height;
            Array.Copy(columnRatios, originalColumnRatios, columnRatios.Length);

            _statusTimer = new DispatcherTimer();
            _statusTimer.Interval = TimeSpan.FromMilliseconds(500);
            _statusTimer.Tick += StatusTimer_Tick;
            _statusTimer.Start();

            DisableCloseButton();

            _queueMonitorSignal = new SemaphoreSlim(0, int.MaxValue);

            Task.Run(MonitorDownloadQueue);

            _lastReportedConcurrentDownloads = DEFAULT_MAX_CONCURRENT_DOWNLOADS;
            _lastReportedDownloadThreads = DEFAULT_MAX_DOWNLOAD_THREADS;
        }

        private async Task MonitorDownloadQueue()
        {
            while (!_queueCts.IsCancellationRequested)
            {
                try
                {
                    await _queueMonitorSignal.WaitAsync(500, _queueCts.Token);

                    while (_activeDownloadCount < DownloadManager.GetMaxConcurrentDownloads())
                    {
                        if (_downloadQueue.TryDequeue(out VideoItem? item))
                        {
                            if (item.StatusFlag != VideoItem.DownloadStatus.Downloading &&
                                item.StatusFlag != VideoItem.DownloadStatus.Completed)
                            {
                                StartDownloadForItem(item);
                            }
                            else
                            {
                                _downloadQueue.Enqueue(item);
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    if (_downloadQueue.IsEmpty && _activeDownloadCount == 0 && _isDownloadStarted)
                    {
                        AppendLog("视频下载已完成");
                        Dispatcher.Invoke(() =>
                        {
                            progressBar.Value = 100;
                        });
                        _isDownloadStarted = false;
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    AppendLog($"队列监控出错: {ex.Message}");
                }
            }
        }

        private void StartDownloadForItem(VideoItem item)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    item.Status = "下载中...";
                    item.StatusColor = Brushes.Orange;
                    item.StatusFlag = VideoItem.DownloadStatus.Downloading;
                    lstResults.Items.Refresh();
                });

                item.DownloadTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_globalCts.Token);
                var token = item.DownloadTokenSource.Token;

                Interlocked.Increment(ref _activeDownloadCount);

                var task = Task.Run(async () =>
                {
                    try
                    {
                        bool result = await DownloadSingleVideoAsync(item, token);
                        return result;
                    }
                    finally
                    {
                        Interlocked.Decrement(ref _activeDownloadCount);

                        if (item.DownloadTokenSource != null)
                        {
                            item.DownloadTokenSource.Dispose();
                            item.DownloadTokenSource = null;
                        }

                        _queueMonitorSignal?.Release();

                        CheckAllDownloadsCompleted();
                    }
                }, token);

                _activeDownloadTasks.Add(task);
            }
            catch (Exception ex)
            {
                AppendLog($"启动下载失败: {ex.Message}");
            }
        }

        private void CheckAllDownloadsCompleted()
        {
            Dispatcher.Invoke(() =>
            {
                bool anyActive = videoItems.Any(item =>
                    item.StatusFlag == VideoItem.DownloadStatus.Downloading ||
                    (item.StatusFlag == VideoItem.DownloadStatus.Pending &&
                     _downloadQueue.Contains(item)));

                bool allCompleted = !anyActive && videoItems.Count > 0 &&
                    videoItems.All(item =>
                        item.StatusFlag == VideoItem.DownloadStatus.Completed ||
                        item.StatusFlag == VideoItem.DownloadStatus.Failed);

                if (allCompleted && _isDownloadStarted)
                {
                    AppendLog("所有视频下载已完成");
                    progressBar.Value = 100;
                    _isDownloadStarted = false;
                }
            });
        }

        private void StatusTimer_Tick(object? sender, EventArgs e)
        {
            UpdateStatusSummary();
        }

        private void UpdateStatusSummary()
        {
            bool anyDownloading = false;
            double totalSpeed = 0;

            foreach (var item in videoItems)
            {
                if (item.StatusFlag == VideoItem.DownloadStatus.Downloading)
                {
                    anyDownloading = true;
                    totalSpeed += item.CurrentSpeedBytesPerSec;
                }
            }

            double percentage = 0;
            if (_totalBytesToDownload > 0)
            {
                percentage = (double)_totalDownloadedBytes / _totalBytesToDownload * 100;
                percentage = Math.Min(percentage, 100);
            }

            Dispatcher.Invoke(() =>
            {
                if (!anyDownloading && _isDownloadStarted)
                {
                    txtStatusSummary.Text = "等待队列中的任务开始下载...";
                }
                else if (!anyDownloading && !_isDownloadStarted && videoItems.Count > 0)
                {
                    txtStatusSummary.Text = "当前没有视频正在下载";
                }
                else if (!anyDownloading)
                {
                    txtStatusSummary.Text = "当前没有视频正在下载，统计信息为空";
                }
                else
                {
                    txtStatusSummary.Text =
                        $"当前下载速度: {FormatTotalSpeed(totalSpeed)} | " +
                        $"已下载: {_completedDownloads}/{videoItems.Count} | " +
                        $"下载完成: {percentage:F1}% | " +
                        $"活动下载数: {_activeDownloadCount}/{DownloadManager.GetMaxConcurrentDownloads()}";
                }
            });
        }

        private string FormatTotalSpeed(double bytesPerSecond)
        {
            if (bytesPerSecond >= 1 << 30)
                return $"{(bytesPerSecond / (1 << 30)):F2} Gb/s";
            if (bytesPerSecond >= 1 << 20)
                return $"{(bytesPerSecond / (1 << 20)):F2} Mb/s";
            if (bytesPerSecond >= 1 << 10)
                return $"{(bytesPerSecond / (1 << 10)):F2} Kb/s";
            return $"{bytesPerSecond:F2} B/s";
        }

        private void DisableCloseButton()
        {
            try
            {
                IntPtr hWnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
                IntPtr hMenu = GetSystemMenu(hWnd, false);
                if (hMenu != IntPtr.Zero)
                {
                    DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"禁用关闭按钮时出错: {ex.Message}");
            }
        }

        private void Window_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_isClosing) return;
            _isClosing = true;

            if (MessageBox.Show("确定要退出程序吗？", "提示",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                e.Cancel = true;
                _isClosing = false;
            }
            else
            {
                CancelAllDownloads();

                Task.Run(async () =>
                {
                    await Task.WhenAll(_activeDownloadTasks.ToArray());
                });
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _isInitialized = true;

            logScrollViewer.ScrollChanged += LogScrollViewer_ScrollChanged;

            sliderMaxConcurrentDownloads.Value = DEFAULT_MAX_CONCURRENT_DOWNLOADS;
            sliderMaxDownloadThreads.Value = DEFAULT_MAX_DOWNLOAD_THREADS;

            DownloadManager.SetMaxConcurrentDownloads(DEFAULT_MAX_CONCURRENT_DOWNLOADS);
            DownloadManager.MaxDownloadThreads = DEFAULT_MAX_DOWNLOAD_THREADS;

            txtMaxConcurrentDownloadsFull.Text = $"当前值: {DEFAULT_MAX_CONCURRENT_DOWNLOADS}";
            txtMaxDownloadThreadsFull.Text = $"当前值: {DEFAULT_MAX_DOWNLOAD_THREADS}";

            UpdateDownloadPathDisplay();

            _totalBytesToDownload = 0;
            _totalDownloadedBytes = 0;
            _completedDownloads = 0;

            lstResults.ItemsSource = null;
            lstResults.ItemsSource = videoItems;

            progressBar.Value = 0;
            ResetLayoutSizes();

            UpdateStatusSummary();

            AppendLog("应用程序初始化完成");
        }

        private void LogScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            _logScrollToEnd = Math.Abs(e.VerticalOffset - (e.ExtentHeight - e.ViewportHeight)) < 1;
        }

        private void MainWindow_Closed(object? sender, EventArgs e)
        {
            _clipboardCts?.Cancel();
            CancelAllDownloads();
            _queueCts.Cancel();

            _queueMonitorSignal?.Dispose();
        }

        private void CancelAllDownloads()
        {
            _globalCts?.Cancel();

            _globalCts?.Dispose();
            _globalCts = new CancellationTokenSource();

            foreach (var item in videoItems)
            {
                if (item.StatusFlag == VideoItem.DownloadStatus.Downloading &&
                    item.DownloadTokenSource != null)
                {
                    item.DownloadTokenSource.Cancel();
                    Dispatcher.Invoke(() =>
                    {
                        item.Status = "已取消";
                        item.StatusColor = Brushes.Orange;
                        lstResults.Items.Refresh();
                    });
                }
            }
        }

        private void UpdateDownloadPathDisplay()
        {
            if (!_isInitialized) return;

            double gridWidth = statusBarGrid.ActualWidth;
            if (gridWidth <= 0) return;

            double pathWidth = txtDownloadPath.ActualWidth;
            double buttonWidth = btnChangePath.ActualWidth;

            var converter = new PathMaxWidthConverter();
            double maxWidth = (double)converter.Convert(
                new object[] { gridWidth, pathWidth, buttonWidth },
                typeof(double),
                null!,
                CultureInfo.CurrentCulture);

            double fullPathWidth = MeasureTextWidth(downloadPath, txtDownloadPath);

            if (fullPathWidth <= maxWidth)
            {
                txtDownloadPath.Text = downloadPath;
            }
            else
            {
                txtDownloadPath.Text = GetEllipsisPath(downloadPath, maxWidth);
            }

            txtDownloadPath.ToolTip = downloadPath;
        }

        private double MeasureTextWidth(string text, TextBlock textBlock)
        {
            FormattedText formattedText = new FormattedText(
                text,
                CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface(textBlock.FontFamily, textBlock.FontStyle, textBlock.FontWeight, textBlock.FontStretch),
                textBlock.FontSize,
                Brushes.Black,
                new NumberSubstitution(),
                VisualTreeHelper.GetDpi(this).PixelsPerDip);

            return formattedText.Width;
        }

        private string GetEllipsisPath(string fullPath, double maxWidth)
        {
            const string ellipsis = "...";
            double ellipsisWidth = MeasureTextWidth(ellipsis, txtDownloadPath);

            int startChars = 10;
            int endChars = 10;
            string candidate = fullPath.Substring(0, startChars) + ellipsis +
                              fullPath.Substring(fullPath.Length - endChars);

            double candidateWidth = MeasureTextWidth(candidate, txtDownloadPath);

            while (candidateWidth > maxWidth && (startChars > 1 || endChars > 1))
            {
                if (startChars > 1) startChars--;
                if (endChars > 1) endChars--;

                candidate = fullPath.Substring(0, startChars) + ellipsis +
                            fullPath.Substring(fullPath.Length - endChars);
                candidateWidth = MeasureTextWidth(candidate, txtDownloadPath);
            }

            return candidate;
        }

        private async void CopyCommand_Click(object sender, RoutedEventArgs e)
        {
            _clipboardCts?.Cancel();
            _clipboardCts = new CancellationTokenSource();
            var token = _clipboardCts.Token;

            try
            {
                btnCopyCommand.IsEnabled = false;
                progressBar.Value = 10;

                bool success = await SafeSetClipboard("getLessonRecordInfo", token);

                if (success)
                {
                    progressBar.Value = 100;
                    AppendLog("筛选文本已成功复制到剪贴板");
                }
                else
                {
                    AppendLog("复制失败，请手动复制命令");
                }
            }
            catch (OperationCanceledException)
            {
                AppendLog("复制操作已取消");
            }
            catch (Exception ex)
            {
                AppendLog($"复制命令失败: {ex.Message}");
            }
            finally
            {
                btnCopyCommand.IsEnabled = true;
            }
        }

        private Task<bool> SafeSetClipboard(string text, CancellationToken token)
        {
            var tcs = new TaskCompletionSource<bool>();

            Thread staThread = new Thread(() =>
            {
                try
                {
                    token.ThrowIfCancellationRequested();
                    int retryCount = 0;
                    bool success = false;

                    while (!success && retryCount < 5)
                    {
                        try
                        {
                            if (NativeMethods.OpenClipboard(IntPtr.Zero))
                            {
                                try
                                {
                                    NativeMethods.EmptyClipboard();
                                    IntPtr hGlobal = Marshal.StringToHGlobalUni(text);
                                    if (NativeMethods.SetClipboardData(13, hGlobal) != IntPtr.Zero)
                                    {
                                        success = true;
                                        tcs.SetResult(true);
                                    }
                                    else
                                    {
                                        Marshal.FreeHGlobal(hGlobal);
                                    }
                                }
                                finally
                                {
                                    NativeMethods.CloseClipboard();
                                }
                            }

                            if (!success)
                            {
                                Thread.Sleep(100);
                                retryCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            if (retryCount >= 4) tcs.SetException(ex);
                            Thread.Sleep(100);
                            retryCount++;
                        }
                    }

                    if (!success) tcs.SetResult(false);
                }
                catch (OperationCanceledException) { tcs.SetCanceled(); }
                catch (Exception ex) { tcs.SetException(ex); }
            });

            staThread.SetApartmentState(ApartmentState.STA);
            staThread.IsBackground = true;
            staThread.Start();

            return tcs.Task;
        }

        private async void Parse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!Clipboard.ContainsText())
                {
                    AppendLog("剪贴板中没有文本内容");
                    return;
                }

                progressBar.Value = 30;

                string clipboardText = Clipboard.GetText();
                var (parseSuccess, duplicateFound) = await ParseContentAsync(clipboardText);

                if (parseSuccess)
                {
                    AppendLog("请求头解析成功");
                    progressBar.Value = 100;
                }
                else if (duplicateFound)
                {
                    progressBar.Value = 0;
                }
                else
                {
                    AppendLog("请求头解析失败");
                    progressBar.Value = 0;
                }
            }
            catch (Exception ex)
            {
                AppendLog($"解析失败: {ex.Message}");
                progressBar.Value = 0;
            }
        }

        private async Task<(bool parseSuccess, bool duplicateFound)> ParseContentAsync(string content)
        {
            List<string> duplicateFiles = new List<string>();
            List<string> duplicateCourses = new List<string>();
            bool parseSuccess = false;
            bool duplicateFound = false;

            try
            {
                JsonDocument doc = JsonDocument.Parse(content);
                JsonElement root = doc.RootElement;

                if (root.TryGetProperty("data", out JsonElement data))
                {
                    string lessonName = data.GetProperty("lessonName").GetString() ?? string.Empty;
                    string safeName = CleanName(lessonName);
                    string outputFile = Path.Combine(downloadPath, $"{safeName}.mp4");

                    if (File.Exists(outputFile))
                    {
                        duplicateFiles.Add($"{safeName}.mp4");
                        duplicateFound = true;
                        AppendLog($"已跳过重复视频项: {lessonName}");
                        return (false, true);
                    }

                    if (videoItems.Any(item => string.Equals(item.Name, lessonName, StringComparison.OrdinalIgnoreCase)))
                    {
                        duplicateCourses.Add(lessonName);
                        duplicateFound = true;
                        AppendLog($"已跳过重复视频项: {lessonName}");
                        return (false, true);
                    }

                    string lastValidUrl = string.Empty;
                    if (data.TryGetProperty("lessonData", out JsonElement lessonData) &&
                        lessonData.TryGetProperty("fileList", out JsonElement fileList) &&
                        fileList.ValueKind == JsonValueKind.Array)
                    {
                        foreach (JsonElement file in fileList.EnumerateArray())
                        {
                            if (file.TryGetProperty("Playset", out JsonElement playset) &&
                                playset.ValueKind == JsonValueKind.Array)
                            {
                                foreach (JsonElement play in playset.EnumerateArray())
                                {
                                    if (play.TryGetProperty("Url", out JsonElement urlElement))
                                    {
                                        string url = urlElement.GetString()?.Replace("\\", "") ?? "";
                                        if (url.Contains(".mp4", StringComparison.OrdinalIgnoreCase))
                                        {
                                            lastValidUrl = url;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    long fileSize = 0;
                    if (!string.IsNullOrEmpty(lastValidUrl))
                    {
                        fileSize = await DownloadManager.GetFileSizeAsync(lastValidUrl);
                    }

                    var newItem = new VideoItem
                    {
                        Name = lessonName,
                        Url = lastValidUrl,
                        FileSize = fileSize,
                        Status = "解析完成",
                        StatusColor = Brushes.Green
                    };

                    Interlocked.Add(ref _totalBytesToDownload, fileSize);
                    videoItems.Add(newItem);
                    UpdateVideoIndexes();

                    Dispatcher.Invoke(() =>
                    {
                        lstResults.ItemsSource = null;
                        lstResults.ItemsSource = videoItems;
                    });

                    if (!string.IsNullOrEmpty(lastValidUrl))
                    {
                        parseSuccess = true;
                    }
                }
                else
                {
                    (parseSuccess, duplicateFound) = await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
                }
            }
            catch (JsonException)
            {
                (parseSuccess, duplicateFound) = await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
            }
            finally
            {
                ShowDuplicateMessage(duplicateFiles, duplicateCourses);
            }

            return (parseSuccess, duplicateFound);
        }

        private async Task<(bool parseSuccess, bool duplicateFound)> UseOriginalLineParsingAsync(string content, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            string[] lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            string currentLessonName = string.Empty;
            string finalUrl = string.Empty;
            bool playsetEncountered = false;
            bool inFileItem = false;
            bool parseSuccess = false;
            bool duplicateFound = false;

            foreach (string line in lines)
            {
                if (line.Contains("lessonName", StringComparison.OrdinalIgnoreCase))
                {
                    currentLessonName = ExtractValue(line, "lessonName");
                }
                else if (line.Contains("{") && !string.IsNullOrEmpty(currentLessonName) && !inFileItem)
                {
                    inFileItem = true;
                    playsetEncountered = false;
                }
                else if (line.Contains("}") && inFileItem && !string.IsNullOrEmpty(currentLessonName))
                {
                    inFileItem = false;
                }

                if (inFileItem && line.Contains("Playset", StringComparison.OrdinalIgnoreCase))
                {
                    playsetEncountered = true;
                }

                if (inFileItem && line.Contains("url", StringComparison.OrdinalIgnoreCase) &&
                   line.Contains("mp4", StringComparison.OrdinalIgnoreCase))
                {
                    string videoUrl = ExtractValue(line, "url").Replace("\\", "");
                    if (playsetEncountered) finalUrl = videoUrl;
                }
            }

            if (!string.IsNullOrEmpty(currentLessonName))
            {
                string safeName = CleanName(currentLessonName);
                string outputFile = Path.Combine(downloadPath, $"{safeName}.mp4");

                if (File.Exists(outputFile))
                {
                    duplicateFiles.Add($"{safeName}.mp4");
                    duplicateFound = true;
                    AppendLog($"已跳过重复视频项: {currentLessonName}");
                    return (false, true);
                }

                if (videoItems.Any(item => string.Equals(item.Name, currentLessonName, StringComparison.OrdinalIgnoreCase)))
                {
                    duplicateCourses.Add(currentLessonName);
                    duplicateFound = true;
                    AppendLog($"已跳过重复视频项: {currentLessonName}");
                    return (false, true);
                }

                long fileSize = 0;
                if (!string.IsNullOrEmpty(finalUrl))
                {
                    fileSize = await DownloadManager.GetFileSizeAsync(finalUrl);
                }

                VideoItem newItem = new VideoItem
                {
                    Name = currentLessonName,
                    Url = finalUrl,
                    FileSize = fileSize,
                    Status = "解析完成",
                    StatusColor = Brushes.Green
                };

                Interlocked.Add(ref _totalBytesToDownload, fileSize);
                videoItems.Add(newItem);
                UpdateVideoIndexes();

                Dispatcher.Invoke(() =>
                {
                    lstResults.ItemsSource = null;
                    lstResults.ItemsSource = videoItems;
                });

                parseSuccess = true;
            }

            return (parseSuccess, duplicateFound);
        }

        private void ShowDuplicateMessage(List<string> duplicateFiles, List<string> duplicateCourses)
        {
            if (duplicateFiles.Count == 0 && duplicateCourses.Count == 0) return;

            StringBuilder message = new StringBuilder();
            if (duplicateFiles.Count > 0)
            {
                message.AppendLine($"跳过 {duplicateFiles.Count} 个已存在的视频项：");
                foreach (var file in duplicateFiles.Take(5)) message.AppendLine($"- {Truncate(file, 50)}");
                if (duplicateFiles.Count > 5) message.AppendLine($"...及其他 {duplicateFiles.Count - 5} 个");
                message.AppendLine();
            }

            if (duplicateCourses.Count > 0)
            {
                message.AppendLine($"跳过 {duplicateCourses.Count} 个重复的请求头：");
                foreach (var course in duplicateCourses.Take(5)) message.AppendLine($"- {Truncate(course, 50)}");
                if (duplicateCourses.Count > 5) message.AppendLine($"...及其他 {duplicateCourses.Count - 5} 个");
            }

            if (message.Length > 0)
            {
                Dispatcher.Invoke(() => MessageBox.Show(message.ToString(), "重复内容提示",
                    MessageBoxButton.OK, MessageBoxImage.Information));
            }
        }

        private void UpdateVideoIndexes()
        {
            for (int i = 0; i < videoItems.Count; i++)
            {
                videoItems[i].DisplayIndex = i + 1;
            }

            if (lstResults.ItemsSource != null)
            {
                Dispatcher.Invoke(() => lstResults.Items.Refresh());
            }
        }

        private string ExtractValue(string jsonLine, string key)
        {
            try
            {
                int keyIndex = jsonLine.IndexOf(key, StringComparison.OrdinalIgnoreCase);
                if (keyIndex < 0) return string.Empty;

                int colonIndex = jsonLine.IndexOf(':', keyIndex + key.Length);
                if (colonIndex < 0) return string.Empty;

                int startIndex = colonIndex + 1;
                while (startIndex < jsonLine.Length && char.IsWhiteSpace(jsonLine[startIndex])) startIndex++;

                int endIndex = startIndex;
                if (startIndex < jsonLine.Length)
                {
                    char startChar = jsonLine[startIndex];
                    char endChar = startChar == '"' ? '"' : ',';
                    endIndex = startChar == '"' ?
                        jsonLine.IndexOf(endChar, startIndex + 1) :
                        jsonLine.IndexOfAny(new[] { ',', '}', ']' }, startIndex);
                }

                if (endIndex < 0) endIndex = jsonLine.Length;

                return jsonLine.Substring(startIndex, endIndex - startIndex)
                    .Trim().Trim('"', '\'', ',', ' ');
            }
            catch
            {
                return string.Empty;
            }
        }

        private async void Download_Click(object sender, RoutedEventArgs e)
        {
            bool allCompleted = videoItems.All(v =>
                v.StatusFlag == VideoItem.DownloadStatus.Completed);

            if (allCompleted)
            {
                AppendLog("没有可下载的视频");
                return;
            }

            var pendingItems = videoItems.Where(item =>
                item.StatusFlag != VideoItem.DownloadStatus.Completed &&
                item.StatusFlag != VideoItem.DownloadStatus.Downloading
            ).ToList();

            if (pendingItems.Count == 0)
            {
                AppendLog("没有可下载的视频");
                return;
            }

            try
            {
                bool anyActiveDownloads = videoItems.Any(item =>
                    item.StatusFlag == VideoItem.DownloadStatus.Downloading
                );

                if (!anyActiveDownloads)
                {
                    AppendLog("开始下载视频...");
                    _isDownloadStarted = true;
                }

                await Task.Run(() => DownloadVideosAsync());
            }
            catch (Exception ex)
            {
                AppendLog($"启动下载失败: {ex.Message}");
            }
        }

        private void DownloadSingle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is VideoItem item)
            {
                try
                {
                    bool allCompleted = videoItems.All(v =>
                        v.StatusFlag == VideoItem.DownloadStatus.Completed);

                    if (allCompleted)
                    {
                        AppendLog("没有可下载的视频");
                        return;
                    }

                    if (item.StatusFlag == VideoItem.DownloadStatus.Completed)
                    {
                        AppendLog("没有可下载的视频");
                        return;
                    }

                    if (item.StatusFlag == VideoItem.DownloadStatus.Downloading)
                    {
                        return;
                    }

                    if (!_isDownloadStarted)
                    {
                        AppendLog("开始下载视频...");
                        _isDownloadStarted = true;
                    }

                    Dispatcher.Invoke(() =>
                    {
                        if (item.StatusFlag != VideoItem.DownloadStatus.Downloading)
                        {
                            item.Status = "等待下载";
                            item.StatusColor = Brushes.Blue;
                            item.StatusFlag = VideoItem.DownloadStatus.Pending;
                        }
                        lstResults.Items.Refresh();
                    });

                    _downloadQueue.Enqueue(item);

                    _queueMonitorSignal?.Release();
                }
                catch (Exception ex)
                {
                    AppendLog($"下载单个视频时出错: {ex.Message}");
                }
            }
        }

        private async Task DownloadVideosAsync()
        {
            List<string> duplicateFiles = new List<string>();

            try
            {
                var videosToDownload = videoItems
                    .Where(item => item.StatusFlag != VideoItem.DownloadStatus.Completed &&
                                   item.StatusFlag != VideoItem.DownloadStatus.Downloading)
                    .ToList();

                if (videosToDownload.Count == 0)
                {
                    return;
                }

                var activeDownloads = new List<VideoItem>();

                foreach (var item in videosToDownload)
                {
                    string safeName = CleanName(item.Name);
                    string outputFile = Path.Combine(downloadPath, $"{safeName}.mp4");

                    if (File.Exists(outputFile))
                    {
                        var fileInfo = new FileInfo(outputFile);
                        if (fileInfo.Length == item.FileSize)
                        {
                            duplicateFiles.Add($"{safeName}.mp4");
                            await Dispatcher.InvokeAsync(() =>
                            {
                                item.Status = "下载完成";
                                item.Progress = 100;
                                item.StatusColor = Brushes.Green;
                                item.StatusFlag = VideoItem.DownloadStatus.Completed;
                                lstResults.Items.Refresh();
                            });

                            await Dispatcher.InvokeAsync(() =>
                            {
                                _completedDownloads++;
                                Interlocked.Add(ref _totalDownloadedBytes, item.FileSize);
                                UpdateStatusSummary();
                            });
                            continue;
                        }
                    }

                    await Dispatcher.InvokeAsync(() =>
                    {
                        if (item.StatusFlag != VideoItem.DownloadStatus.Partial &&
                            item.StatusFlag != VideoItem.DownloadStatus.Downloading)
                        {
                            item.Status = "等待下载";
                            item.StatusColor = Brushes.Blue;
                            item.StatusFlag = VideoItem.DownloadStatus.Pending;
                        }
                        lstResults.Items.Refresh();
                    });

                    activeDownloads.Add(item);
                }

                foreach (var item in activeDownloads)
                {
                    _downloadQueue.Enqueue(item);
                    _queueMonitorSignal?.Release();
                }
            }
            catch (Exception ex)
            {
                AppendLog($"下载过程中出错: {ex.Message}");
            }
        }

        private async Task<bool> DownloadSingleVideoAsync(VideoItem item, CancellationToken token)
        {
            string safeName = CleanName(item.Name);
            string outputFile = Path.Combine(downloadPath, $"{safeName}.mp4");
            int retryCount = 0;
            const int maxRetries = 3;

            using var timeoutCts = new CancellationTokenSource(TimeSpan.FromHours(6));
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(token, timeoutCts.Token);
            var linkedToken = linkedCts.Token;

            while (retryCount <= maxRetries)
            {
                try
                {
                    if (string.IsNullOrEmpty(item.Url))
                    {
                        Dispatcher.Invoke(() =>
                        {
                            item.Status = "下载失败: 无URL";
                            item.StatusColor = Brushes.Red;
                            item.StatusFlag = VideoItem.DownloadStatus.Failed;
                            lstResults.Items.Refresh();
                        });
                        return false;
                    }

                    if (File.Exists(outputFile))
                    {
                        var fileInfo = new FileInfo(outputFile);
                        long fileSize = fileInfo.Length;

                        if (item.FileSize > 0 && fileSize == item.FileSize)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                item.Status = "下载完成";
                                item.Progress = 100;
                                item.StatusColor = Brushes.Green;
                                item.StatusFlag = VideoItem.DownloadStatus.Completed;
                                lstResults.Items.Refresh();
                            });

                            Dispatcher.Invoke(() => _completedDownloads++);
                            return true;
                        }
                        else if (fileSize > 0 && fileSize < item.FileSize)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                item.DownloadedBytes = fileSize;
                                item.Progress = (int)(fileSize * 100 / item.FileSize);
                                item.Status = "继续下载";
                                item.StatusColor = Brushes.Blue;
                                item.StatusFlag = VideoItem.DownloadStatus.Partial;
                                lstResults.Items.Refresh();
                            });
                        }
                    }

                    item.LastReportedBytes = item.DownloadedBytes;

                    void ProgressCallback(long downloaded, long total, double speed, TimeSpan remaining)
                    {
                        var now = DateTime.Now;
                        if ((now - _lastUiUpdate).TotalMilliseconds < UI_UPDATE_THROTTLE_MS) return;
                        _lastUiUpdate = now;

                        Dispatcher.BeginInvoke(new Action(() =>
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
                                long delta = downloaded - item.LastReportedBytes;
                                item.LastReportedBytes = downloaded;

                                Interlocked.Add(ref _totalDownloadedBytes, delta);

                                item.DownloadedBytes = downloaded;
                                item.TotalBytes = total;

                                if (total > 0)
                                {
                                    item.Progress = (int)(downloaded * 100 / total);
                                }

                                item.CurrentSpeedBytesPerSec = speed;
                                item.RemainingTime = remaining;
                            }
                        }), DispatcherPriority.Background);
                    }

                    int currentThreads = DownloadManager.MaxDownloadThreads;

                    bool result = await DownloadManager.DownloadFileAsync(
                        item.Url,
                        outputFile,
                        ProgressCallback,
                        currentThreads,
                        linkedToken
                    );

                    if (result)
                    {
                        await Task.Delay(500);

                        var fileInfo = new FileInfo(outputFile);
                        if (item.FileSize > 0 && fileInfo.Length != item.FileSize)
                        {
                            throw new Exception($"文件验证失败: 实际大小 {fileInfo.Length}, 预期大小 {item.FileSize}");
                        }

                        Dispatcher.Invoke(() =>
                        {
                            item.Progress = 100;
                            item.DownloadedBytes = item.TotalBytes;
                            item.CurrentSpeedBytesPerSec = 0;
                            item.RemainingTime = TimeSpan.Zero;
                            item.Status = "下载完成";
                            item.StatusColor = Brushes.Green;
                            item.StatusFlag = VideoItem.DownloadStatus.Completed;
                            lstResults.Items.Refresh();
                        });

                        Dispatcher.Invoke(() => _completedDownloads++);
                        return true;
                    }
                    else
                    {
                        if (File.Exists(outputFile) && new FileInfo(outputFile).Length > 0)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                item.Status = "已暂停";
                                item.StatusColor = Brushes.Orange;
                                item.StatusFlag = VideoItem.DownloadStatus.Partial;
                                lstResults.Items.Refresh();
                            });
                            return false;
                        }
                        else
                        {
                            throw new Exception("下载失败");
                        }
                    }
                }
                catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested)
                {
                    AppendLog($"下载超时: {item.Name}");
                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        item.Status = "下载超时";
                        item.StatusColor = Brushes.Red;
                        item.StatusFlag = VideoItem.DownloadStatus.Failed;
                        lstResults.Items.Refresh();
                    }), DispatcherPriority.Background);
                    return false;
                }
                catch (OperationCanceledException)
                {
                    if (File.Exists(outputFile) && new FileInfo(outputFile).Length > 0)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            item.Status = "已暂停";
                            item.Phase = "暂停";
                            item.StatusFlag = VideoItem.DownloadStatus.Partial;
                            lstResults.Items.Refresh();
                        });
                    }
                    else
                    {
                        Dispatcher.Invoke(() =>
                        {
                            item.Status = "已取消";
                            item.Phase = "取消";
                            item.StatusFlag = VideoItem.DownloadStatus.Pending;
                            lstResults.Items.Refresh();
                        });
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    retryCount++;

                    if (retryCount > maxRetries)
                    {
                        AppendLog($"下载失败: {ex.Message}");
                        Dispatcher.Invoke(() =>
                        {
                            item.Status = "下载失败";
                            item.StatusColor = Brushes.Red;
                            item.StatusFlag = VideoItem.DownloadStatus.Failed;
                            lstResults.Items.Refresh();
                        });
                        return false;
                    }
                    else
                    {
                        int delay = (int)Math.Pow(2, retryCount) * 1000;
                        AppendLog($"下载失败，将在 {delay}ms 后重试 ({retryCount}/{maxRetries}): {ex.Message}");
                        await Task.Delay(delay, linkedToken);
                    }
                }
                finally
                {
                    if (item.DownloadTokenSource != null)
                    {
                        item.DownloadTokenSource.Dispose();
                        item.DownloadTokenSource = null;
                    }
                }
            }

            return false;
        }

        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Directory.Exists(downloadPath))
                {
                    Process.Start("explorer.exe", downloadPath);
                    AppendLog("正在打开下载目录...");

                    progressBar.Value = 100;

                    var resetTimer = new DispatcherTimer();
                    resetTimer.Interval = TimeSpan.FromSeconds(1);
                    resetTimer.Tick += (s, args) => {
                        resetTimer.Stop();
                        progressBar.Value = 0;
                    };
                    resetTimer.Start();
                }
                else
                {
                    AppendLog("下载目录不存在");
                }
            }
            catch (Exception ex)
            {
                AppendLog($"无法打开目录: {ex.Message}");
            }
        }

        private void ChangeDownloadPath_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new OpenFolderDialog
                {
                    Title = "选择下载目录",
                    InitialDirectory = Directory.Exists(downloadPath) ? downloadPath : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                };

                if (dialog.ShowDialog() == true)
                {
                    downloadPath = dialog.FolderName;
                    UpdateDownloadPathDisplay();
                    AppendLog($"下载路径已更新为: {downloadPath}");
                }
            }
            catch (Exception ex)
            {
                AppendLog($"更改目录失败: {ex.Message}");
            }
        }

        private async void CopyUrl_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string url)
            {
                try
                {
                    bool success = await SafeSetClipboard(url, CancellationToken.None);
                    if (success) AppendLog("视频URL已复制到剪贴板");
                    else AppendLog("复制URL失败，请重试");
                }
                catch (Exception ex)
                {
                    AppendLog($"复制URL失败: {ex.Message}");
                }
            }
        }

        private void DeleteItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is VideoItem item)
            {
                try
                {
                    string logPrefix = $"删除视频项: '{Truncate(item.Name, 50)}' - ";

                    if (item.StatusFlag == VideoItem.DownloadStatus.Downloading)
                    {
                        AppendLog($"{logPrefix}正在下载中，需要取消下载");

                        if (MessageBox.Show($"视频 '{Truncate(item.Name, 50)}' 正在下载中，删除将取消下载。\n是否继续删除？",
                                "警告", MessageBoxButton.OKCancel, MessageBoxImage.Warning) != MessageBoxResult.OK)
                        {
                            AppendLog($"{logPrefix}用户取消了删除操作");
                            return;
                        }

                        if (item.DownloadTokenSource != null)
                        {
                            AppendLog($"{logPrefix}正在取消下载任务");
                            item.DownloadTokenSource.Cancel();
                            AppendLog($"{logPrefix}下载任务已取消");
                        }
                    }

                    if (item.StatusFlag == VideoItem.DownloadStatus.Completed)
                    {
                        _completedDownloads--;
                        AppendLog($"{logPrefix}从已完成计数中移除");
                    }

                    if (item.FileSize > 0)
                    {
                        Interlocked.Add(ref _totalBytesToDownload, -item.FileSize);

                        if (item.StatusFlag == VideoItem.DownloadStatus.Completed)
                        {
                            Interlocked.Add(ref _totalDownloadedBytes, -item.FileSize);
                            AppendLog($"{logPrefix}从总下载字节中移除完整文件大小");
                        }
                        else
                        {
                            Interlocked.Add(ref _totalDownloadedBytes, -item.DownloadedBytes);
                        }
                    }

                    videoItems.Remove(item);

                    UpdateVideoIndexes();

                    UpdateStatusSummary();
                    AppendLog($"{logPrefix} 成功");
                }
                catch (Exception ex)
                {
                    string errorMsg = $"删除视频项 '{Truncate(item.Name, 30)}' 失败: {ex.Message}";
                    AppendLog(errorMsg);
                    MessageBox.Show(errorMsg, "删除错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void InitializeTool_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                bool hasPendingItems = videoItems.Any(item =>
                    item.StatusFlag != VideoItem.DownloadStatus.Completed &&
                    item.StatusFlag != VideoItem.DownloadStatus.Failed);

                bool hasActiveDownloads = videoItems.Any(item =>
                    item.StatusFlag == VideoItem.DownloadStatus.Downloading);

                if (hasPendingItems || hasActiveDownloads)
                {
                    string message = "当前有视频未完成下载或正在下载中，继续初始化将取消所有下载任务并清除视频列表。\n是否继续初始化？";

                    if (MessageBox.Show(message, "警告",
                        MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                    {
                        AppendLog("用户取消了初始化操作");
                        return;
                    }
                }

                bool anyActiveDownloads = videoItems.Any(item =>
                    item.StatusFlag == VideoItem.DownloadStatus.Downloading
                );

                if (anyActiveDownloads)
                {
                    CancelAllDownloads();

                    AppendLog("正在取消下载任务...");
                    var waitTask = Task.WhenAll(_activeDownloadTasks.ToArray());

                    var completedTask = await Task.WhenAny(waitTask, Task.Delay(2000));

                    if (completedTask != waitTask)
                    {
                        AppendLog("部分任务仍在取消中，继续初始化");
                    }
                }

                AppendLog("开始初始化解析工具...");

                Dispatcher.Invoke(() =>
                {
                    this.Width = 1280;
                    this.Height = 720;
                    this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                });

                _globalCts?.Dispose();
                _globalCts = new CancellationTokenSource();
                _activeDownloadTasks = new ConcurrentBag<Task>();
                _downloadQueue = new ConcurrentQueue<VideoItem>();
                _activeDownloadCount = 0;
                _isDownloadStarted = false;

                Dispatcher.Invoke(() =>
                {
                    sliderMaxConcurrentDownloads.Value = DEFAULT_MAX_CONCURRENT_DOWNLOADS;
                    sliderMaxDownloadThreads.Value = DEFAULT_MAX_DOWNLOAD_THREADS;

                    DownloadManager.SetMaxConcurrentDownloads(DEFAULT_MAX_CONCURRENT_DOWNLOADS);
                    DownloadManager.MaxDownloadThreads = DEFAULT_MAX_DOWNLOAD_THREADS;

                    txtMaxConcurrentDownloadsFull.Text = $"当前值: {DEFAULT_MAX_CONCURRENT_DOWNLOADS}";
                    txtMaxDownloadThreadsFull.Text = $"当前值: {DEFAULT_MAX_DOWNLOAD_THREADS}";

                    txtLog.Clear();
                    videoItems.Clear();

                    downloadPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "下载目录");
                    Directory.CreateDirectory(downloadPath);
                    UpdateDownloadPathDisplay();

                    _totalBytesToDownload = 0;
                    _totalDownloadedBytes = 0;
                    _completedDownloads = 0;

                    lstResults.ItemsSource = null;
                    lstResults.ItemsSource = videoItems;

                    progressBar.Value = 0;
                    ResetLayoutSizes();

                    UpdateStatusSummary();
                });

                AppendLog("解析工具初始化完成");
            }
            catch (Exception ex)
            {
                AppendLog($"初始化工具时出错: {ex.Message}");
            }
        }

        private void ResetLayoutSizes()
        {
            Dispatcher.Invoke(() =>
            {
                PanelColumn.Width = originalPanelWidth;
                ResultColumn.Width = originalResultWidth;
                MainContentRow.Height = originalMainContentHeight;
                LogRow.Height = originalLogHeight;

                Array.Copy(originalColumnRatios, columnRatios, originalColumnRatios.Length);
                userAdjustedColumns = false;

                if (lstResults.View is GridView gridView && gridView.Columns.Count == 4)
                {
                    double totalWidth = lstResults.ActualWidth - SystemParameters.VerticalScrollBarWidth;
                    if (totalWidth > 0)
                    {
                        gridView.Columns[0].Width = totalWidth * columnRatios[0];
                        gridView.Columns[1].Width = totalWidth * columnRatios[1];
                        gridView.Columns[2].Width = totalWidth * columnRatios[2];
                        gridView.Columns[3].Width = totalWidth * columnRatios[3];
                    }
                }
            });
        }

        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_isInitialized) return;

            if (sender == sliderMaxConcurrentDownloads)
            {
                int value = (int)e.NewValue;
                txtMaxConcurrentDownloadsFull.Text = $"当前值: {value}";
                DownloadManager.SetMaxConcurrentDownloads(value);
            }
            else if (sender == sliderMaxDownloadThreads)
            {
                int value = (int)e.NewValue;
                txtMaxDownloadThreadsFull.Text = $"当前值: {value}";
                DownloadManager.MaxDownloadThreads = value;
            }
        }

        private void Slider_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (sender == sliderMaxConcurrentDownloads)
            {
                int value = (int)sliderMaxConcurrentDownloads.Value;
                if (value != _lastReportedConcurrentDownloads)
                {
                    AppendLog($"最大并发下载数已更新为: {value}");
                    _lastReportedConcurrentDownloads = value;
                }
            }
            else if (sender == sliderMaxDownloadThreads)
            {
                int value = (int)sliderMaxDownloadThreads.Value;
                if (value != _lastReportedDownloadThreads)
                {
                    AppendLog($"最大下载线程数已更新为: {value}");
                    _lastReportedDownloadThreads = value;
                }
            }
        }

        private void GridViewColumnHeader_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is GridViewColumnHeader header)
            {
                var thumb = FindVisualChild<Thumb>(header);
                if (thumb != null)
                {
                    thumb.DragDelta += (s, args) =>
                    {
                        userAdjustedColumns = true;
                        if (lstResults.View is GridView gridView && gridView.Columns.Count == 4)
                        {
                            double totalWidth = gridView.Columns[0].ActualWidth +
                                               gridView.Columns[1].ActualWidth +
                                               gridView.Columns[2].ActualWidth +
                                               gridView.Columns[3].ActualWidth;

                            if (totalWidth > 0)
                            {
                                columnRatios[0] = gridView.Columns[0].ActualWidth / totalWidth;
                                columnRatios[1] = gridView.Columns[1].ActualWidth / totalWidth;
                                columnRatios[2] = gridView.Columns[2].ActualWidth / totalWidth;
                                columnRatios[3] = gridView.Columns[3].ActualWidth / totalWidth;
                            }
                        }
                    };
                }
            }
        }

        private T? FindVisualChild<T>(DependencyObject? parent) where T : DependencyObject
        {
            if (parent == null) return null;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T result)
                    return result;

                var descendant = FindVisualChild<T>(child);
                if (descendant != null)
                    return descendant;
            }
            return null;
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateLayoutBindings();

            UpdateDownloadPathDisplay();
        }

        private void UpdateLayoutBindings()
        {
            foreach (var child in FindVisualChildren<FrameworkElement>(this))
            {
                if (child.GetBindingExpression(TextBlock.FontSizeProperty) != null)
                {
                    child.GetBindingExpression(TextBlock.FontSizeProperty)?.UpdateTarget();
                }

                if (child.GetBindingExpression(Control.FontSizeProperty) != null)
                {
                    child.GetBindingExpression(Control.FontSizeProperty)?.UpdateTarget();
                }

                if (child.GetBindingExpression(FrameworkElement.HeightProperty) != null)
                {
                    child.GetBindingExpression(FrameworkElement.HeightProperty)?.UpdateTarget();
                }
            }
        }

        private IEnumerable<T> FindVisualChildren<T>(DependencyObject? depObj) where T : DependencyObject
        {
            if (depObj == null) yield break;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                var child = VisualTreeHelper.GetChild(depObj, i);
                if (child != null && child is T t)
                {
                    yield return t;
                }

                foreach (var childOfChild in FindVisualChildren<T>(child))
                {
                    yield return childOfChild;
                }
            }
        }

        private void ListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (lstResults.View is GridView gridView && gridView.Columns.Count == 4)
            {
                double totalWidth = lstResults.ActualWidth - SystemParameters.VerticalScrollBarWidth;
                if (totalWidth > 0)
                {
                    if (userAdjustedColumns)
                    {
                        gridView.Columns[0].Width = totalWidth * columnRatios[0];
                        gridView.Columns[1].Width = totalWidth * columnRatios[1];
                        gridView.Columns[2].Width = totalWidth * columnRatios[2];
                        gridView.Columns[3].Width = totalWidth * columnRatios[3];
                    }
                    else if (lastTotalWidth != totalWidth)
                    {
                        gridView.Columns[0].Width = totalWidth * columnRatios[0];
                        gridView.Columns[1].Width = totalWidth * columnRatios[1];
                        gridView.Columns[2].Width = totalWidth * columnRatios[2];
                        gridView.Columns[3].Width = totalWidth * columnRatios[3];
                        lastTotalWidth = totalWidth;
                    }
                }
            }
        }

        private void AppendLog(string message)
        {
            if (!_isInitialized) return;

            string cleanMessage = message.Replace("\r", "").Replace("\n", " ");
            string logEntry = $"{DateTime.Now:HH:mm:ss} - {cleanMessage}";

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (txtLog == null) return;

                if (!string.IsNullOrEmpty(txtLog.Text) && !txtLog.Text.EndsWith(Environment.NewLine))
                {
                    txtLog.AppendText(Environment.NewLine);
                }
                txtLog.AppendText(logEntry);

                if (_logScrollToEnd)
                {
                    txtLog.CaretIndex = txtLog.Text.Length;
                    logScrollViewer.ScrollToEnd();
                }
            }), DispatcherPriority.Background);

            try
            {
                string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "VideoDownloader.log");
                File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"写入日志文件失败: {ex.Message}");
            }
        }

        private string Truncate(string text, int maxLength)
        {
            return string.IsNullOrEmpty(text) || text.Length <= maxLength ?
                text : text.Substring(0, maxLength) + "...";
        }

        private string CleanName(string name)
        {
            if (string.IsNullOrEmpty(name)) return name;

            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                name = name.Replace(c.ToString(), "");
            }

            return name.Replace(":", "").Replace("?", "").Replace("*", "")
                     .Replace("|", "").Replace("<", "").Replace(">", "").Replace("\"", "");
        }
    }

    /// <summary>
    /// 下载管理器类，用于管理视频下载任务，包括并发控制、文件大小获取和下载执行
    /// </summary>
    public static class DownloadManager
    {
        /// <summary>
        /// HTTP客户端，用于发送下载请求
        /// </summary>
        private static readonly HttpClient _httpClient;
        /// <summary>
        /// 活动下载任务字典，用于跟踪正在进行的下载任务
        /// </summary>
        private static readonly ConcurrentDictionary<string, CancellationTokenSource> _activeDownloads = new ConcurrentDictionary<string, CancellationTokenSource>();
        /// <summary>
        /// 最大并发下载数，默认值为5
        /// </summary>
        private static int _maxConcurrentDownloads = 5;
        /// <summary>
        /// 最大下载线程数，默认值为32
        /// </summary>
        private static int _maxDownloadThreads = 32;

        /// <summary>
        /// 并发控制信号量，用于限制同时进行的下载任务数
        /// </summary>
        private static SemaphoreSlim _concurrencySemaphore;
        /// <summary>
        /// 锁对象，用于保护_maxConcurrentDownloads的修改
        /// </summary>
        private static readonly object _lock = new object();

        /// <summary>
        /// 静态构造函数，初始化HTTP客户端和相关配置
        /// </summary>
        static DownloadManager()
        {
            // 初始化并发控制信号量
            _concurrencySemaphore = new SemaphoreSlim(_maxConcurrentDownloads, _maxConcurrentDownloads);

            // 配置HTTP客户端处理器
            var handler = new HttpClientHandler
            {
                UseProxy = false,           // 不使用代理
                Proxy = null,               // 代理为空
                MaxConnectionsPerServer = 100, // 每个服务器最大连接数为100
                UseDefaultCredentials = false, // 不使用默认凭据
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate // 启用自动解压缩
            };

            // 初始化HTTP客户端
            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromHours(6), // 超时时间为6小时
                DefaultRequestHeaders =
                {
                    ConnectionClose = false, // 不关闭连接
                    CacheControl = new CacheControlHeaderValue { NoCache = true } // 不使用缓存
                }
            };

            // 配置ServicePointManager
            ServicePointManager.DefaultConnectionLimit = 100; // 默认连接限制为100
            ServicePointManager.ReusePort = true; // 启用端口重用
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13; // 启用TLS 1.2和1.3
            ServicePointManager.MaxServicePointIdleTime = 10000; // 最大服务点空闲时间为10秒
        }

        /// <summary>
        /// 设置最大并发下载数
        /// </summary>
        /// <param name="max">最大并发下载数</param>
        public static void SetMaxConcurrentDownloads(int max)
        {
            lock (_lock)
            {
                // 如果新值与当前值相同，直接返回
                if (max == _maxConcurrentDownloads)
                    return;

                int oldMax = _maxConcurrentDownloads;
                _maxConcurrentDownloads = max;

                if (max > oldMax)
                {
                    // 如果新值大于旧值，释放额外的信号量
                    _concurrencySemaphore.Release(max - oldMax);
                }
                else
                {
                    // 如果新值小于旧值，尝试获取并释放多余的信号量
                    int diff = oldMax - max;
                    for (int i = 0; i < diff; i++)
                    {
                        if (_concurrencySemaphore.CurrentCount > 0)
                        {
                            _concurrencySemaphore.Wait(0);
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 获取当前最大并发下载数
        /// </summary>
        /// <returns>当前最大并发下载数</returns>
        public static int GetMaxConcurrentDownloads() => _maxConcurrentDownloads;

        /// <summary>
        /// 最大下载线程数属性，限制在1-256之间
        /// </summary>
        public static int MaxDownloadThreads
        {
            get => _maxDownloadThreads;
            set => _maxDownloadThreads = Math.Clamp(value, 1, 256);
        }

        /// <summary>
        /// 获取文件大小
        /// </summary>
        /// <param name="url">文件URL</param>
        /// <returns>文件大小（字节），如果获取失败则返回0</returns>
        public static async Task<long> GetFileSizeAsync(string url)
        {
            try
            {
                // 发送HEAD请求获取文件大小
                using var response = await GetHeadAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    // 返回Content-Length头的值，如果不存在则返回0
                    return response.Content.Headers.ContentLength ?? 0;
                }
            }
            catch
            {
                // 忽略所有异常，返回0
            }
            return 0;
        }

        public static async Task<HttpResponseMessage> GetHeadAsync(string url, CancellationToken cancellationToken = default)
        {
            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Head, url))
                {
                    request.Headers.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
                    request.Headers.Accept.ParseAdd("video/mp4, */*");
                    request.Headers.Referrer = new Uri("https://www.eeo.cn/");

                    return await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
                }
            }
            catch
            {
                return new HttpResponseMessage(HttpStatusCode.BadRequest);
            }
        }

        public static async Task<bool> DownloadFileAsync(string url, string filePath,
            Action<long, long, double, TimeSpan> progressCallback,
            int threads,
            CancellationToken cancellationToken = default)
        {
            const int maxRetries = 5;
            int retryCount = 0;
            long totalBytesRead = 0;
            long? totalBytes = null;
            long startPosition = 0;
            long totalBytesValue = 0;
            bool isFinalUpdateSent = false;

            while (retryCount <= maxRetries)
            {
                try
                {
                    if (string.IsNullOrEmpty(url))
                    {
                        throw new ArgumentException("URL不能为空");
                    }

                    if (!totalBytes.HasValue)
                    {
                        using var headResponse = await GetHeadAsync(url, cancellationToken);
                        if (headResponse.IsSuccessStatusCode)
                        {
                            totalBytes = headResponse.Content.Headers.ContentLength;
                        }
                    }

                    totalBytesValue = totalBytes ?? 0;

                    string? directory = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    using var request = new HttpRequestMessage(HttpMethod.Get, url);

                    request.Headers.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
                    request.Headers.Accept.ParseAdd("video/mp4, */*");
                    request.Headers.Referrer = new Uri("https://www.eeo.cn/");

                    if (startPosition > 0)
                    {
                        request.Headers.Range = new RangeHeaderValue(startPosition, null);
                    }

                    using var response = await _httpClient.SendAsync(
                        request,
                        HttpCompletionOption.ResponseHeadersRead,
                        cancellationToken);

                    if (!response.IsSuccessStatusCode)
                    {
                        if (response.StatusCode != HttpStatusCode.PartialContent)
                        {
                            throw new HttpRequestException($"服务器返回错误状态: {response.StatusCode}");
                        }
                    }

                    if (response.StatusCode == HttpStatusCode.PartialContent)
                    {
                        var contentLength = response.Content.Headers.ContentLength;
                        if (contentLength.HasValue)
                        {
                            totalBytes = startPosition + contentLength.Value;
                            totalBytesValue = totalBytes.Value;
                        }
                    }
                    else if (!totalBytes.HasValue)
                    {
                        totalBytes = response.Content.Headers.ContentLength;
                        totalBytesValue = totalBytes ?? 0;
                    }

                    using var contentStream = await response.Content.ReadAsStreamAsync();

                    int bufferSize = 1024 * 1024;
                    var buffer = new byte[bufferSize];
                    var lastBytesRead = startPosition;
                    var lastUpdateTime = DateTime.Now;

                    using (var fileStream = new FileStream(
                        filePath,
                        startPosition > 0 ? FileMode.Append : FileMode.Create,
                        FileAccess.Write,
                        FileShare.None,
                        bufferSize,
                        FileOptions.Asynchronous))
                    {
                        while (true)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            var bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length, cancellationToken);
                            if (bytesRead == 0) break;

                            await fileStream.WriteAsync(buffer, 0, bytesRead, cancellationToken);

                            totalBytesRead += bytesRead;

                            var elapsed = (DateTime.Now - lastUpdateTime).TotalMilliseconds;
                            if (elapsed > 200)
                            {
                                var speed = (totalBytesRead - lastBytesRead) / (elapsed / 1000);
                                lastBytesRead = totalBytesRead;
                                lastUpdateTime = DateTime.Now;

                                var remainingTime = totalBytesValue > 0
                                    ? TimeSpan.FromSeconds((totalBytesValue - totalBytesRead) / (speed > 0 ? speed : 1))
                                    : TimeSpan.MaxValue;

                                progressCallback?.Invoke(totalBytesRead, totalBytesValue, speed, remainingTime);
                            }
                        }

                        await fileStream.FlushAsync();
                    }

                    progressCallback?.Invoke(totalBytesValue, totalBytesValue, 0, TimeSpan.Zero);
                    isFinalUpdateSent = true;

                    await Task.Delay(100);

                    return true;
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (HttpRequestException ex) when (ex.StatusCode.HasValue &&
                    (int)ex.StatusCode.Value >= 500 &&
                    retryCount < maxRetries)
                {
                    retryCount++;
                    await Task.Delay(ExponentialBackoff(retryCount), cancellationToken);
                }
                catch (IOException ioEx) when (
                    ioEx is FileNotFoundException ||
                    ioEx is DirectoryNotFoundException ||
                    ioEx.HResult == -2147024894)
                {
                    throw;
                }
                catch (Exception ex) when (retryCount < maxRetries)
                {
                    Debug.WriteLine($"下载失败，正在重试({retryCount}/{maxRetries}): {ex.Message}");
                    retryCount++;
                    await Task.Delay(ExponentialBackoff(retryCount), cancellationToken);
                }
                finally
                {
                    if (!isFinalUpdateSent)
                    {
                        progressCallback?.Invoke(totalBytesRead, totalBytesValue, 0, TimeSpan.Zero);
                        isFinalUpdateSent = true;
                    }
                }
            }

            return false;
        }

        private static int ExponentialBackoff(int retryCount)
        {
            return (int)Math.Min(1000 * Math.Pow(2, retryCount), 60000);
        }

        public static void CancelDownload(string url)
        {
            if (_activeDownloads.TryRemove(url, out var cts))
            {
                cts.Cancel();
            }
        }

        public static async Task<bool> DownloadWithConcurrencyControl(
            string url,
            string filePath,
            Action<long, long, double, TimeSpan> progressCallback,
            CancellationToken cancellationToken,
            int threads,
            Action? onDownloadStarted = null)
        {
            var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            _activeDownloads[url] = cts;

            try
            {
                await _concurrencySemaphore.WaitAsync(cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();

                onDownloadStarted?.Invoke();

                return await DownloadFileAsync(url, filePath, progressCallback, threads, cts.Token);
            }
            finally
            {
                try
                {
                    _concurrencySemaphore.Release();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"释放信号量时出错: {ex.Message}");
                }

                _activeDownloads.TryRemove(url, out _);
                cts.Dispose();
            }
        }
    }

    internal static class NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EmptyClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
    }
}