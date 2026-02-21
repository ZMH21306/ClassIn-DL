using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using Microsoft.Win32;
using System.ComponentModel;
using System.Globalization;
using System.Windows.Threading;
using System.Collections.ObjectModel;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Net.Http;
using System.Net;
using System.Collections.Concurrent;
using System.Net.Http.Headers;

namespace VideoDownloader
{
    // 转换器：将数字索引格式化为两位数显示（如1变为"01"）
    public class IndexFormatConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is int index ? index.ToString("D2") : "00";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // 转换器：根据列表高度动态计算每个列表项的高度
    public class HeightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double height)
            {
                double itemHeight = (height - 5) / 10.0;
                return Math.Clamp(itemHeight, 30, 80);
            }
            return 40.0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // 转换器：用于按钮悬停/按下状态时颜色变暗效果
    public class DarkenColorConverter : IValueConverter
    {
        public double DarkenFactor { get; set; } = 0.3;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is SolidColorBrush brush)
            {
                Color color = brush.Color;
                byte r = (byte)(color.R * (1 - DarkenFactor));
                byte g = (byte)(color.G * (1 - DarkenFactor));
                byte b = (byte)(color.B * (1 - DarkenFactor));
                return new SolidColorBrush(Color.FromRgb(r, g, b));
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public partial class MainWindow : Window
    {
        // 默认设置值
        private const int DEFAULT_MAX_CONCURRENT_DOWNLOADS = 5;
        private const int DEFAULT_MAX_DOWNLOAD_THREADS = 32;

        private string downloadPath = string.Empty;
        private ObservableCollection<VideoItem> videoItems = new ObservableCollection<VideoItem>();
        private CancellationTokenSource? _clipboardCts;
        private bool _isClosing = false;

        // 列表视图列宽比例
        private double[] columnRatios = { 0.05, 0.3, 0.4, 0.25 };
        private bool userAdjustedColumns = false;
        private double lastTotalWidth = 0;

        // 初始布局尺寸
        private double[] originalColumnRatios = { 0.05, 0.3, 0.4, 0.25 };
        private GridLength originalPanelWidth;
        private GridLength originalResultWidth;
        private GridLength originalMainContentHeight;
        private GridLength originalLogHeight;

        // UI状态标志
        private bool _logScrollToEnd = true;
        private bool _isInitialized = false;

        // 下载状态跟踪
        private long _totalBytesToDownload = 0;
        private long _totalDownloadedBytes = 0;
        private int _completedDownloads = 0;
        private DispatcherTimer _statusTimer;

        // 下载控制
        private CancellationTokenSource _globalCts = new CancellationTokenSource();
        private ConcurrentBag<Task> _activeDownloadTasks = new ConcurrentBag<Task>();

        // Win32 API声明
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

        // 表示单个视频下载任务的数据结构
        public class VideoItem : INotifyPropertyChanged
        {
            // 视频下载状态枚举
            public enum DownloadStatus
            {
                Pending,      // 等待下载
                Downloading,   // 下载中
                Completed,     // 已完成
                Failed,        // 失败
                Partial        // 部分下载
            }

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

            // 指示当前视频是否正在下载中
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

            // 视频文件的总大小
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

            // 下载状态文本描述
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

            // 下载进度百分比
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
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            // 已下载的字节数
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
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            // 视频文件的总字节数
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
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            // 当前下载速度（字节/秒）
            private double _currentSpeedBytesPerSec;
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

            // 预计剩余下载时间
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
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            // 当前下载阶段描述
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

            // 上次报告的已下载字节数（用于计算增量）
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

            public CancellationTokenSource? DownloadTokenSource { get; set; }

            // 组合多个属性生成的状态显示文本
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

            // 格式化文件大小为易读的字符串（自动选择合适单位）
            public static string FormatSize(long bytes)
            {
                if (bytes >= 1 << 30) return $"{bytes / (1 << 30):F2} GB";
                if (bytes >= 1 << 20) return $"{bytes / (1 << 20):F2} MB";
                if (bytes >= 1 << 10) return $"{bytes / (1 << 10):F2} KB";
                return $"{bytes} B";
            }

            // 格式化下载速度为易读的字符串（自动选择合适单位）
            public static string FormatSpeed(double bytesPerSecond)
            {
                if (bytesPerSecond >= 1 << 30) return $"{bytesPerSecond / (1 << 30):F2} GB/s";
                if (bytesPerSecond >= 1 << 20) return $"{bytesPerSecond / (1 << 20):F2} MB/s";
                if (bytesPerSecond >= 1 << 10) return $"{bytesPerSecond / (1 << 10):F2} KB/s";
                return $"{bytesPerSecond:F2} B/s";
            }

            // 格式化时间为HH:MM:SS格式
            public static string FormatTime(TimeSpan time)
            {
                if (time == TimeSpan.MaxValue) return "未知";
                return $"{time.Hours:D2}:{time.Minutes:D2}:{time.Seconds:D2}";
            }

            public Brush StatusColor { get; set; } = Brushes.Gray;

            public event PropertyChangedEventHandler? PropertyChanged;

            protected virtual void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public MainWindow()
        {
            InitializeComponent();

            // 初始化下载目录
            downloadPath = Path.Combine(Environment.CurrentDirectory, "下载目录");
            Directory.CreateDirectory(downloadPath);

            this.Closed += MainWindow_Closed;

            // 保存初始布局尺寸
            originalPanelWidth = PanelColumn.Width;
            originalResultWidth = ResultColumn.Width;
            originalMainContentHeight = MainContentRow.Height;
            originalLogHeight = LogRow.Height;
            Array.Copy(columnRatios, originalColumnRatios, columnRatios.Length);

            // 初始化状态更新定时器（500ms更新一次）
            _statusTimer = new DispatcherTimer();
            _statusTimer.Interval = TimeSpan.FromMilliseconds(500);
            _statusTimer.Tick += StatusTimer_Tick;
            _statusTimer.Start();

            // 禁用窗口关闭按钮
            DisableCloseButton();
        }

        // 状态定时器回调：定期更新下载状态摘要信息
        private void StatusTimer_Tick(object? sender, EventArgs e)
        {
            UpdateStatusSummary();
        }

        // 更新状态摘要信息：显示下载速度、进度等统计信息
        private void UpdateStatusSummary()
        {
            // 检查是否有视频正在下载
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

            // 计算下载完成百分比
            double percentage = 0;
            if (_totalBytesToDownload > 0)
            {
                percentage = (double)_totalDownloadedBytes / _totalBytesToDownload * 100;
                percentage = Math.Min(percentage, 100);  // 确保不超过100%
            }

            // 更新状态摘要文本
            Dispatcher.Invoke(() =>
            {
                if (!anyDownloading)
                {
                    txtStatusSummary.Text = "当前没有视频正在下载，统计信息为空";
                }
                else
                {
                    txtStatusSummary.Text =
                        $"当前下载速度: {FormatTotalSpeed(totalSpeed)} | " +
                        $"已下载: {_completedDownloads}/{videoItems.Count} | " +
                        $"下载完成: {percentage:F1}%";
                }
            });
        }

        // 格式化总下载速度为易读字符串（自动选择合适单位）
        private string FormatTotalSpeed(double bytesPerSecond)
        {
            if (bytesPerSecond >= 1 << 30) // 1GB/s
                return $"{(bytesPerSecond / (1 << 30)):F2} Gb/s";
            if (bytesPerSecond >= 1 << 20) // 1MB/s
                return $"{(bytesPerSecond / (1 << 20)):F2} Mb/s";
            if (bytesPerSecond >= 1 << 10) // 1KB/s
                return $"{(bytesPerSecond / (1 << 10)):F2} Kb/s";
            return $"{bytesPerSecond:F2} B/s";
        }

        // 禁用窗口关闭按钮（防止下载过程中意外关闭）
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

        // 窗口关闭处理：确认关闭并取消所有下载
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
                // 取消所有下载任务
                CancelAllDownloads();

                // 非阻塞等待任务完成取消
                Task.Run(async () => {
                    await Task.WhenAll(_activeDownloadTasks.ToArray());
                });
            }
        }

        // 窗口加载完成事件处理
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _isInitialized = true;

            // 注册日志滚动事件
            logScrollViewer.ScrollChanged += LogScrollViewer_ScrollChanged;

            // 初始化设置值
            sliderMaxConcurrentDownloads.Value = DEFAULT_MAX_CONCURRENT_DOWNLOADS;
            sliderMaxDownloadThreads.Value = DEFAULT_MAX_DOWNLOAD_THREADS;

            // 应用初始设置
            DownloadManager.SetMaxConcurrentDownloads(DEFAULT_MAX_CONCURRENT_DOWNLOADS);
            DownloadManager.MaxDownloadThreads = DEFAULT_MAX_DOWNLOAD_THREADS;

            // 更新UI显示
            txtMaxConcurrentDownloadsFull.Text = $"当前值: {DEFAULT_MAX_CONCURRENT_DOWNLOADS}";
            txtMaxDownloadThreadsFull.Text = $"当前值: {DEFAULT_MAX_DOWNLOAD_THREADS}";

            // 更新下载路径显示
            UpdateDownloadPathDisplay();

            // 初始状态更新
            UpdateStatusSummary();

            // 记录初始化日志
            AppendLog("应用程序已初始化");
            txtStatus.Text = "应用程序已就绪";
            txtCurrentAction.Text = "请开始操作";
        }

        // 日志滚动事件处理：检测是否滚动到底部
        private void LogScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            _logScrollToEnd = Math.Abs(e.VerticalOffset - (e.ExtentHeight - e.ViewportHeight)) < 1;
        }

        // 窗口关闭事件处理：取消所有操作
        private void MainWindow_Closed(object? sender, EventArgs e)
        {
            _clipboardCts?.Cancel();
            CancelAllDownloads();
        }

        // 取消所有下载任务
        private void CancelAllDownloads()
        {
            // 使用全局取消令牌取消所有任务
            _globalCts?.Cancel();

            // 重置全局取消令牌
            _globalCts?.Dispose();
            _globalCts = new CancellationTokenSource();

            // 取消所有视频项的下载
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

        // 更新下载路径显示
        private void UpdateDownloadPathDisplay()
        {
            if (!_isInitialized) return;

            txtDownloadPath.Text = downloadPath;
            txtDownloadPath.ToolTip = downloadPath;
        }

        // "复制筛选关键词"按钮点击处理
        private async void CopyCommand_Click(object sender, RoutedEventArgs e)
        {
            _clipboardCts?.Cancel();
            _clipboardCts = new CancellationTokenSource();
            var token = _clipboardCts.Token;

            try
            {
                btnCopyCommand.IsEnabled = false;
                progressBar.Value = 10;

                // 尝试将命令复制到剪贴板
                bool success = await SafeSetClipboard("getLessonRecordInfo", token);

                if (success)
                {
                    Dispatcher.Invoke(() =>
                    {
                        txtCurrentAction.Text = "已将 'getLessonRecordInfo' 复制到剪贴板";
                        txtStatus.Text = "请到浏览器开发者工具中执行此命令";
                        txtStatus.Foreground = Brushes.Green;
                    });
                    progressBar.Value = 100;
                    AppendLog("已成功将筛选项复制到剪贴板");
                }
                else
                {
                    UpdateStatus("复制失败: 请手动复制命令", Brushes.Orange);
                }
            }
            catch (OperationCanceledException)
            {
                UpdateStatus("复制操作已取消", Brushes.Orange);
            }
            catch (Exception ex)
            {
                UpdateStatus($"复制命令失败: {ex.Message}", Brushes.Red);
            }
            finally
            {
                btnCopyCommand.IsEnabled = true;
            }
        }

        // 安全设置剪贴板内容（使用STA线程操作）
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

                    // 重试机制（最多5次）
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

        // 更新状态文本和颜色
        private void UpdateStatus(string message, Brush color)
        {
            if (!_isInitialized) return;

            Dispatcher.Invoke(() =>
            {
                txtStatus.Text = message;
                txtStatus.Foreground = color;
            });
        }

        // "解析剪贴板"按钮点击处理
        private async void Parse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!Clipboard.ContainsText())
                {
                    UpdateStatus("剪贴板中没有文本内容", Brushes.Red);
                    return;
                }

                txtCurrentAction.Text = "正在解析剪贴板内容...";
                progressBar.Value = 30;

                // 获取并解析剪贴板内容
                string clipboardText = Clipboard.GetText();
                await ParseContentAsync(clipboardText);
            }
            catch (Exception ex)
            {
                UpdateStatus($"解析失败: {ex.Message}", Brushes.Red);
            }
        }

        // 解析剪贴板内容（支持JSON和旧格式）
        private async Task ParseContentAsync(string content)
        {
            List<string> duplicateFiles = new List<string>();
            List<string> duplicateCourses = new List<string>();

            try
            {
                JsonDocument doc = JsonDocument.Parse(content);
                JsonElement root = doc.RootElement;

                // 尝试从JSON中提取视频信息
                if (root.TryGetProperty("data", out JsonElement data))
                {
                    string lessonName = data.GetProperty("lessonName").GetString() ?? string.Empty;
                    string safeName = CleanName(lessonName);
                    string outputFile = Path.Combine(downloadPath, $"{safeName}.mp4");

                    // 检查重复文件
                    if (File.Exists(outputFile))
                    {
                        duplicateFiles.Add($"{safeName}.mp4");
                        return;
                    }

                    // 检查重复课程
                    if (videoItems.Any(item => string.Equals(item.Name, lessonName, StringComparison.OrdinalIgnoreCase)))
                    {
                        duplicateCourses.Add(lessonName);
                        return;
                    }

                    string lastValidUrl = string.Empty;
                    // 提取视频URL
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

                    // 获取文件大小
                    long fileSize = 0;
                    if (!string.IsNullOrEmpty(lastValidUrl))
                    {
                        fileSize = await DownloadManager.GetFileSizeAsync(lastValidUrl);
                    }

                    // 创建新视频项
                    var newItem = new VideoItem
                    {
                        Name = lessonName,
                        Url = lastValidUrl,
                        FileSize = fileSize,
                        Status = "解析完成",
                        StatusColor = Brushes.Green
                    };

                    // 更新全局状态
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
                        UpdateStatus($"成功解析视频", Brushes.Green);
                    }
                    else
                    {
                        UpdateStatus("未找到视频信息", Brushes.Orange);
                    }
                }
                else
                {
                    // 如果JSON解析失败，尝试行解析
                    await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
                }

                progressBar.Value = 100;
                txtCurrentAction.Text = "解析完成";
                UpdateStatusSummary();
            }
            catch (JsonException)
            {
                // JSON解析异常时使用行解析
                await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
            }
            finally
            {
                ShowDuplicateMessage(duplicateFiles, duplicateCourses);
            }
        }

        // 使用行解析方法处理内容（兼容旧格式）
        private async Task UseOriginalLineParsingAsync(string content, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            string[] lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            string currentLessonName = string.Empty;
            string finalUrl = string.Empty;
            bool playsetEncountered = false;
            bool inFileItem = false;

            // 行解析逻辑
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

                // 检查重复文件
                if (File.Exists(outputFile))
                {
                    duplicateFiles.Add($"{safeName}.mp4");
                    return;
                }

                // 检查重复课程
                if (videoItems.Any(item => string.Equals(item.Name, currentLessonName, StringComparison.OrdinalIgnoreCase)))
                {
                    duplicateCourses.Add(currentLessonName);
                    return;
                }

                // 获取文件大小
                long fileSize = 0;
                if (!string.IsNullOrEmpty(finalUrl))
                {
                    fileSize = await DownloadManager.GetFileSizeAsync(finalUrl);
                }

                // 创建新视频项
                VideoItem newItem = new VideoItem
                {
                    Name = currentLessonName,
                    Url = finalUrl,
                    FileSize = fileSize,
                    Status = "解析完成",
                    StatusColor = Brushes.Green
                };

                // 更新全局状态
                Interlocked.Add(ref _totalBytesToDownload, fileSize);
                videoItems.Add(newItem);
                UpdateVideoIndexes();

                Dispatcher.Invoke(() =>
                {
                    lstResults.ItemsSource = null;
                    lstResults.ItemsSource = videoItems;
                });
            }
            else
            {
                UpdateStatus("未找到课程名称", Brushes.Orange);
            }

            progressBar.Value = 100;
            txtCurrentAction.Text = "解析完成";
            UpdateStatusSummary();
        }

        // 显示重复文件/课程消息
        private void ShowDuplicateMessage(List<string> duplicateFiles, List<string> duplicateCourses)
        {
            if (duplicateFiles.Count == 0 && duplicateCourses.Count == 0) return;

            StringBuilder message = new StringBuilder();
            if (duplicateFiles.Count > 0)
            {
                message.AppendLine($"跳过 {duplicateFiles.Count} 个已存在的视频文件：");
                foreach (var file in duplicateFiles.Take(5)) message.AppendLine($"- {Truncate(file, 50)}");
                if (duplicateFiles.Count > 5) message.AppendLine($"...及其他 {duplicateFiles.Count - 5} 个");
                message.AppendLine();
            }

            if (duplicateCourses.Count > 0)
            {
                message.AppendLine($"跳过 {duplicateCourses.Count} 个重复课程：");
                foreach (var course in duplicateCourses.Take(5)) message.AppendLine($"- {Truncate(course, 50)}");
                if (duplicateCourses.Count > 5) message.AppendLine($"...及其他 {duplicateCourses.Count - 5} 极速");
            }

            if (message.Length > 0)
            {
                Dispatcher.Invoke(() => MessageBox.Show(message.ToString(), "重复内容提示",
                    MessageBoxButton.OK, MessageBoxImage.Information));
            }
        }

        // 更新视频项索引（序号）
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

        // 从JSON行中提取值
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

        // "下载全部视频"按钮点击处理
        private void Download_Click(object sender, RoutedEventArgs e)
        {
            if (videoItems.Count == 0)
            {
                UpdateStatus("没有可下载的视频", Brushes.Orange);
                return;
            }

            try
            {
                txtCurrentAction.Text = "正在启动下载...";
                progressBar.Value = 50;
                AppendLog("开始下载视频...");

                // 重置全局状态
                _totalDownloadedBytes = 0;
                _completedDownloads = 0;

                // 启动下载任务
                Task.Run(() => DownloadVideosAsync());
            }
            catch (Exception ex)
            {
                UpdateStatus($"启动下载失败: {ex.Message}", Brushes.Red);
            }
        }

        // 单个视频下载按钮点击处理
        private async void DownloadSingle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is VideoItem item)
            {
                try
                {
                    txtCurrentAction.Text = $"正在下载: {Truncate(item.Name, 20)}";
                    AppendLog($"开始下载单个视频: {item.Name}");

                    await DownloadSingleVideoAsync(item, new CancellationTokenSource().Token);
                }
                catch (Exception ex)
                {
                    AppendLog($"下载单个视频时出错: {ex.Message}");
                }
            }
        }

        // 下载所有视频任务（异步）
        private async Task DownloadVideosAsync()
        {
            List<string> duplicateFiles = new List<string>();

            try
            {
                Dispatcher.Invoke(() => txtCurrentAction.Text = "下载进行中...");

                // 获取所有未完成的视频项
                var videosToDownload = videoItems
                    .Where(item => item.StatusFlag != VideoItem.DownloadStatus.Completed)
                    .ToList();

                if (videosToDownload.Count == 0)
                {
                    AppendLog("没有需要下载的视频");
                    Dispatcher.Invoke(() => txtCurrentAction.Text = "没有需要下载的任务");
                    return;
                }

                int successCount = 0;
                int failedCount = 0;
                int skippedCount = 0;
                var downloadTasks = new List<Task>();
                var activeDownloads = new List<VideoItem>();

                // 准备下载任务
                foreach (var item in videosToDownload)
                {
                    string safeName = CleanName(item.Name);
                    string outputFile = Path.Combine(downloadPath, $"{safeName}.mp4");

                    // 检查已完成文件
                    if (File.Exists(outputFile))
                    {
                        var fileInfo = new FileInfo(outputFile);
                        if (fileInfo.Length == item.FileSize)
                        {
                            duplicateFiles.Add($"{safeName}.mp4");
                            skippedCount++;
                            Dispatcher.Invoke(() =>
                            {
                                item.Status = "下载完成";
                                item.Progress = 100;
                                item.StatusColor = Brushes.Green;
                                item.StatusFlag = VideoItem.DownloadStatus.Completed;
                                lstResults.Items.Refresh();
                            });

                            // 更新全局状态
                            Dispatcher.Invoke(() =>
                            {
                                _completedDownloads++;
                                Interlocked.Add(ref _totalDownloadedBytes, item.FileSize);
                                UpdateStatusSummary();
                            });
                            continue;
                        }
                    }

                    // 跳过正在下载的项目
                    if (item.StatusFlag == VideoItem.DownloadStatus.Downloading)
                    {
                        continue;
                    }

                    // 更新等待状态
                    Dispatcher.Invoke(() =>
                    {
                        if (item.StatusFlag != VideoItem.DownloadStatus.Partial)
                        {
                            item.Status = "等待下载";
                            item.StatusColor = Brushes.Blue;
                            item.StatusFlag = VideoItem.DownloadStatus.Pending;
                        }
                        lstResults.Items.Refresh();
                    });

                    activeDownloads.Add(item);
                }

                // 启动下载任务
                foreach (var item in activeDownloads)
                {
                    string safeName = CleanName(item.Name);
                    string outputFile = Path.Combine(downloadPath, $"{safeName}.mp4");

                    if (File.Exists(outputFile)) continue;

                    // 创建下载令牌
                    item.DownloadTokenSource = CancellationTokenSource.CreateLinkedTokenSource(_globalCts.Token);
                    var token = item.DownloadTokenSource.Token;

                    var task = Task.Run(async () =>
                    {
                        try
                        {
                            bool result = await DownloadSingleVideoAsync(item, token);
                            if (result) Interlocked.Increment(ref successCount);
                            else Interlocked.Increment(ref failedCount);
                        }
                        catch (OperationCanceledException)
                        {
                            // 取消操作不计数
                        }
                        finally
                        {
                            // 释放资源
                            if (item.DownloadTokenSource != null)
                            {
                                item.DownloadTokenSource.Dispose();
                                item.DownloadTokenSource = null;
                            }
                        }
                    }, token);

                    _activeDownloadTasks.Add(task);
                    downloadTasks.Add(task);
                }

                // 等待所有任务完成
                if (downloadTasks.Count > 0)
                {
                    await Task.WhenAll(downloadTasks);
                }

                AppendLog($"下载完成: 成功 {successCount} 个, 失败 {failedCount} 个, 跳过 {skippedCount} 个");

                Dispatcher.Invoke(() =>
                {
                    progressBar.Value = 100;
                    txtCurrentAction.Text = "下载完成";

                    if (failedCount == 0)
                        UpdateStatus($"下载完成: {successCount}个成功, {skippedCount}个跳过", Brushes.Green);
                    else
                        UpdateStatus($"下载完成: {successCount}个成功, {failedCount}个失败, {skippedCount}个跳过", Brushes.Orange);
                });

                ShowDuplicateFilesMessage(duplicateFiles);
            }
            catch (Exception ex)
            {
                AppendLog($"下载过程中出错: {ex.Message}");
                Dispatcher.Invoke(() => UpdateStatus($"下载失败: {ex.Message}", Brushes.Red));
            }
            finally
            {
                // 清理完成的任务
                foreach (var task in _activeDownloadTasks.ToList())
                {
                    if (task.IsCompleted)
                    {
                        _activeDownloadTasks.TryTake(out _);
                    }
                }
            }
        }

        // 显示重复文件消息
        private void ShowDuplicateFilesMessage(List<string> duplicateFiles)
        {
            if (duplicateFiles.Count == 0) return;

            StringBuilder message = new StringBuilder();
            if (duplicateFiles.Count > 0)
            {
                message.AppendLine($"跳过 {duplicateFiles.Count} 个已存在的视频文件：");
                foreach (var file in duplicateFiles.Take(5)) message.AppendLine($"- {Truncate(file, 50)}");
                if (duplicateFiles.Count > 5) message.AppendLine($"...及其他 {duplicateFiles.Count - 5} 个");
            }

            if (message.Length > 0)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show(message.ToString(),
                        "文件已存在",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                });
            }
        }

        // 下载单个视频（支持断点续传） - 关键修复：增强下载稳定性
        private async Task<bool> DownloadSingleVideoAsync(VideoItem item, CancellationToken token)
        {
            string safeName = CleanName(item.Name);
            string outputFile = Path.Combine(downloadPath, $"{safeName}.mp4");
            int retryCount = 0;
            const int maxRetries = 3; // 最大重试次数

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

                    // 检查文件是否已存在
                    if (File.Exists(outputFile))
                    {
                        var fileInfo = new FileInfo(outputFile);
                        long fileSize = fileInfo.Length;

                        // 检查文件是否完整
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

                            // 更新全局状态
                            Dispatcher.Invoke(() => _completedDownloads++);
                            return true;
                        }
                        // 处理部分下载
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

                    // 初始化进度跟踪
                    item.LastReportedBytes = item.DownloadedBytes;

                    // 进度回调函数
                    void ProgressCallback(long downloaded, long total, double speed, TimeSpan remaining)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            // 下载完成处理
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
                                // 更新下载进度
                                long delta = downloaded - item.LastReportedBytes;
                                item.LastReportedBytes = downloaded;

                                // 更新全局进度
                                Interlocked.Add(ref _totalDownloadedBytes, delta);

                                // 更新当前项
                                item.DownloadedBytes = downloaded;
                                item.TotalBytes = total;

                                if (total > 0)
                                {
                                    item.Progress = (int)(downloaded * 100 / total);
                                }

                                item.CurrentSpeedBytesPerSec = speed;
                                item.RemainingTime = remaining;
                            }
                        });
                    }

                    // 获取当前线程设置
                    int currentThreads = DownloadManager.MaxDownloadThreads;

                    // 更新等待状态
                    Dispatcher.Invoke(() =>
                    {
                        if (item.StatusFlag != VideoItem.DownloadStatus.Partial)
                        {
                            item.Status = "等待下载";
                            item.StatusColor = Brushes.Blue;
                            item.StatusFlag = VideoItem.DownloadStatus.Pending;
                        }
                        lstResults.Items.Refresh();
                    });

                    // 开始下载
                    bool result = await DownloadManager.DownloadWithConcurrencyControl(
                        item.Url,
                        outputFile,
                        ProgressCallback,
                        token,
                        currentThreads,
                        () => // 下载开始回调
                        {
                            Dispatcher.Invoke(() =>
                            {
                                item.Status = "下载中...";
                                item.StatusColor = Brushes.Orange;
                                item.StatusFlag = VideoItem.DownloadStatus.Downloading;
                                lstResults.Items.Refresh();
                            });
                        }
                    );

                    // 处理下载结果
                    if (result)
                    {
                        // 等待文件写入完成
                        await Task.Delay(500);

                        // 验证文件大小
                        var fileInfo = new FileInfo(outputFile);
                        if (item.FileSize > 0 && fileInfo.Length != item.FileSize)
                        {
                            throw new Exception($"文件验证失败: 实际大小 {fileInfo.Length}, 预期大小 {item.FileSize}");
                        }

                        // 更新完成状态
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

                        // 更新全局状态
                        Dispatcher.Invoke(() => _completedDownloads++);
                        return true;
                    }
                    else
                    {
                        // 部分下载处理
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
                catch (OperationCanceledException)
                {
                    // 取消操作处理
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
                        // 最终失败处理
                        if (ex.Message.Contains("文件验证失败"))
                        {
                            try { File.Delete(outputFile); } catch { }
                            AppendLog($"下载失败: {item.Name} - {ex.Message}");
                            Dispatcher.Invoke(() =>
                            {
                                item.Status = "下载失败: 文件不完整";
                                item.StatusColor = Brushes.Red;
                                item.StatusFlag = VideoItem.DownloadStatus.Failed;
                                lstResults.Items.Refresh();
                            });
                        }
                        else
                        {
                            AppendLog($"下载失败: {ex.Message}");
                            Dispatcher.Invoke(() =>
                            {
                                item.Status = "下载失败";
                                item.StatusColor = Brushes.Red;
                                item.StatusFlag = VideoItem.DownloadStatus.Failed;
                                lstResults.Items.Refresh();
                            });
                        }
                        return false;
                    }
                    else
                    {
                        // 重试前等待
                        int delay = (int)Math.Pow(2, retryCount) * 1000; // 指数退避
                        AppendLog($"下载失败，将在 {delay}ms 后重试 ({retryCount}/{maxRetries}): {ex.Message}");
                        await Task.Delay(delay, token);
                    }
                }
                finally
                {
                    // 确保释放下载令牌
                    if (item.DownloadTokenSource != null)
                    {
                        item.DownloadTokenSource.Dispose();
                        item.DownloadTokenSource = null;
                    }
                }
            }

            return false;
        }

        // "打开下载目录"按钮点击处理
        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Directory.Exists(downloadPath))
                {
                    Process.Start("explorer.exe", downloadPath);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"无法打开目录: {ex.Message}", Brushes.Red);
            }
        }

        // "更改下载路径"按钮点击处理
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
                    UpdateStatus($"下载路径已更新: {downloadPath}", Brushes.Green);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"更改目录失败: {ex.Message}", Brushes.Red);
            }
        }

        // "复制URL"按钮点击处理
        private async void CopyUrl_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string url)
            {
                try
                {
                    bool success = await SafeSetClipboard(url, CancellationToken.None);
                    if (success) UpdateStatus("视频URL已复制到剪贴板", Brushes.Green);
                    else UpdateStatus("复制URL失败，请重试", Brushes.Orange);
                }
                catch (Exception ex)
                {
                    UpdateStatus($"复制URL失败: {ex.Message}", Brushes.Red);
                }
            }
        }

        // "删除项"按钮点击处理
        private void DeleteItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is VideoItem item)
            {
                try
                {
                    // 处理正在下载的任务
                    if (item.StatusFlag == VideoItem.DownloadStatus.Downloading)
                    {
                        if (MessageBox.Show($"视频 '{Truncate(item.Name, 50)}' 正在下载中，删除将取消下载。\n是否继续删除？",
                            "警告", MessageBoxButton.OKCancel, MessageBoxImage.Warning) != MessageBoxResult.OK)
                        {
                            return;
                        }

                        // 取消下载
                        if (item.DownloadTokenSource != null)
                        {
                            item.DownloadTokenSource.Cancel();
                        }
                    }

                    // 更新全局状态
                    if (item.StatusFlag == VideoItem.DownloadStatus.Completed)
                    {
                        _completedDownloads--;
                    }

                    // 更新字节统计
                    if (item.FileSize > 0)
                    {
                        Interlocked.Add(ref _totalBytesToDownload, -item.FileSize);

                        if (item.StatusFlag == VideoItem.DownloadStatus.Completed)
                        {
                            Interlocked.Add(ref _totalDownloadedBytes, -item.FileSize);
                        }
                        else
                        {
                            Interlocked.Add(ref _totalDownloadedBytes, -item.DownloadedBytes);
                        }
                    }

                    // 移除项
                    videoItems.Remove(item);
                    UpdateVideoIndexes();

                    // 更新状态
                    UpdateStatusSummary();
                }
                catch (Exception ex)
                {
                    AppendLog($"删除视频项失败: {ex.Message}");
                }
            }
        }

        // "初始化解析工具"按钮点击处理
        private async void InitializeTool_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 检查是否有活动下载
                bool anyActiveDownloads = videoItems.Any(item =>
                    item.StatusFlag == VideoItem.DownloadStatus.Downloading);

                if (anyActiveDownloads)
                {
                    if (MessageBox.Show("当前有视频正在下载中，继续初始化将取消所有下载任务。\n是否继续初始化？",
                            "警告", MessageBoxButton.OKCancel, MessageBoxImage.Warning) != MessageBoxResult.OK)
                    {
                        AppendLog("用户取消了初始化操作");
                        return;
                    }

                    // 取消所有下载
                    CancelAllDownloads();

                    // 等待任务取消
                    Dispatcher.Invoke(() => txtCurrentAction.Text = "正在取消下载任务...");
                    var waitTask = Task.WhenAll(_activeDownloadTasks.ToArray());
                    var completedTask = await Task.WhenAny(waitTask, Task.Delay(2000));

                    if (completedTask != waitTask)
                    {
                        AppendLog("部分任务仍在取消中，继续初始化");
                    }
                }

                AppendLog("开始初始化解析工具...");

                // 重置全局状态
                _globalCts?.Dispose();
                _globalCts = new CancellationTokenSource();
                _activeDownloadTasks = new ConcurrentBag<Task>();

                Dispatcher.Invoke(() =>
                {
                    // 重置设置
                    sliderMaxConcurrentDownloads.Value = DEFAULT_MAX_CONCURRENT_DOWNLOADS;
                    sliderMaxDownloadThreads.Value = DEFAULT_MAX_DOWNLOAD_THREADS;

                    DownloadManager.SetMaxConcurrentDownloads(DEFAULT_MAX_CONCURRENT_DOWNLOADS);
                    DownloadManager.MaxDownloadThreads = DEFAULT_MAX_DOWNLOAD_THREADS;

                    txtMaxConcurrentDownloadsFull.Text = $"当前值: {DEFAULT_MAX_CONCURRENT_DOWNLOADS}";
                    txtMaxDownloadThreadsFull.Text = $"当前值: {DEFAULT_MAX_DOWNLOAD_THREADS}";

                    // 清除数据和日志
                    txtLog.Clear();
                    videoItems.Clear();

                    // 重置统计
                    _totalBytesToDownload = 0;
                    _totalDownloadedBytes = 0;
                    _completedDownloads = 0;

                    // 重置列表
                    lstResults.ItemsSource = null;
                    lstResults.ItemsSource = videoItems;

                    // 重置UI状态
                    progressBar.Value = 0;
                    txtStatus.Text = "准备就绪...";
                    txtStatus.Foreground = Brushes.Black;
                    txtCurrentAction.Text = "就绪";
                    ResetLayoutSizes();

                    // 更新状态
                    UpdateStatusSummary();
                });

                AppendLog("解析工具初始化完成");
            }
            catch (Exception ex)
            {
                AppendLog($"初始化工具时出错: {ex.Message}");
            }
        }

        // 重置布局尺寸到初始状态
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

                // 重置列表列宽
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

        // 滑块值改变处理：更新设置值
        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_isInitialized) return;

            if (sender == sliderMaxConcurrentDownloads)
            {
                int value = (int)e.NewValue;
                txtMaxConcurrentDownloadsFull.Text = $"当前值: {value}";
                DownloadManager.SetMaxConcurrentDownloads(value);
                UpdateStatus($"最大同时下载数量已更新: {value}", Brushes.Green);
            }
            else if (sender == sliderMaxDownloadThreads)
            {
                int value = (int)e.NewValue;
                txtMaxDownloadThreadsFull.Text = $"当前值: {value}";
                DownloadManager.MaxDownloadThreads = value;
                UpdateStatus($"最大下载线程数量已更新: {value}", Brushes.Green);
            }
        }

        // 列表列头加载处理：添加宽度调整事件
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

        // 在可视化树中查找特定类型的子元素
        private T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
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

        // 列表大小改变处理：调整列宽比例
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

        // 添加日志条目到日志文本框
        private void AppendLog(string message)
        {
            if (!_isInitialized) return;

            Dispatcher.Invoke(() =>
            {
                if (txtLog == null) return;

                // 清理消息并添加时间戳
                message = message.Replace("\r", "").Replace("\n", " ");
                string logEntry = $"{DateTime.Now:HH:mm:ss} - {message}";

                // 添加换行符（如果需要）
                if (!string.IsNullOrEmpty(txtLog.Text) && !txtLog.Text.EndsWith(Environment.NewLine))
                {
                    txtLog.AppendText(Environment.NewLine);
                }
                txtLog.AppendText(logEntry);

                // 自动滚动到底部
                if (_logScrollToEnd)
                {
                    txtLog.CaretIndex = txtLog.Text.Length;
                    logScrollViewer.ScrollToEnd();
                }
            });
        }

        // 截断字符串（添加省略号）
        private string Truncate(string text, int maxLength)
        {
            return string.IsNullOrEmpty(text) || text.Length <= maxLength ?
                text : text.Substring(0, maxLength) + "...";
        }

        // 清理名称（移除非法字符）
        private string CleanName(string name)
        {
            if (string.IsNullOrEmpty(name)) return name;

            // 移除非法文件名字符
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                name = name.Replace(c.ToString(), "");
            }

            // 移除额外特殊字符
            return name.Replace(":", "").Replace("?", "").Replace("*", "")
                     .Replace("|", "").Replace("<", "").Replace(">", "").Replace("\"", "");
        }
    }

    // 下载管理器：处理视频下载的核心逻辑 - 关键修复：增强下载稳定性
    public static class DownloadManager
    {
        private static readonly HttpClient _httpClient;  // HTTP客户端
        private static readonly ConcurrentDictionary<string, CancellationTokenSource> _activeDownloads = new ConcurrentDictionary<string, CancellationTokenSource>();
        private static int _maxConcurrentDownloads = 5;  // 最大并发下载数
        private static int _maxDownloadThreads = 32;     // 最大下载线程数

        // 并发控制信号量
        private static SemaphoreSlim _concurrencySemaphore;
        private static readonly object _lock = new object();

        static DownloadManager()
        {
            // 初始化信号量
            _concurrencySemaphore = new SemaphoreSlim(_maxConcurrentDownloads, _maxConcurrentDownloads);

            // 配置HTTP客户端
            var handler = new HttpClientHandler
            {
                UseProxy = false,
                Proxy = null,
                MaxConnectionsPerServer = 100,
                UseDefaultCredentials = false,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromHours(6),
                DefaultRequestHeaders =
                {
                    ConnectionClose = false,
                    CacheControl = new CacheControlHeaderValue { NoCache = true }
                }
            };

            // 配置网络连接
            ServicePointManager.DefaultConnectionLimit = 100;
            ServicePointManager.ReusePort = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
            ServicePointManager.MaxServicePointIdleTime = 10000; // 10秒空闲时间
        }

        // 设置最大并发下载数
        public static void SetMaxConcurrentDownloads(int max)
        {
            lock (_lock)
            {
                if (max == _maxConcurrentDownloads)
                    return;

                int oldMax = _maxConcurrentDownloads;
                _maxConcurrentDownloads = max;

                // 调整信号量大小
                if (max > oldMax)
                {
                    _concurrencySemaphore.Release(max - oldMax);
                }
                else
                {
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

        // 获取最大并发下载数
        public static int GetMaxConcurrentDownloads() => _maxConcurrentDownloads;

        // 最大下载线程数属性
        public static int MaxDownloadThreads
        {
            get => _maxDownloadThreads;
            set => _maxDownloadThreads = Math.Clamp(value, 1, 256);
        }

        // 获取远程文件大小
        public static async Task<long> GetFileSizeAsync(string url)
        {
            try
            {
                using var response = await GetHeadAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    return response.Content.Headers.ContentLength ?? 0;
                }
            }
            catch
            {
                // 忽略错误
            }
            return 0;
        }

        // 发送HEAD请求获取文件信息
        public static async Task<HttpResponseMessage> GetHeadAsync(string url, CancellationToken cancellationToken = default)
        {
            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Head, url))
                {
                    // 添加浏览器模拟头
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

        // 下载文件（支持断点续传和进度报告） - 关键修复：增强稳定性
        public static async Task<bool> DownloadFileAsync(string url, string filePath,
            Action<long, long, double, TimeSpan> progressCallback,
            int threads,
            CancellationToken cancellationToken = default)
        {
            const int maxRetries = 5;  // 最大重试次数
            int retryCount = 0;
            long totalBytesRead = 0;
            long? totalBytes = null;
            long startPosition = 0;
            long totalBytesValue = 0;  // 修复：移出 try 块
            bool isFinalUpdateSent = false;

            while (retryCount <= maxRetries)
            {
                try
                {
                    if (string.IsNullOrEmpty(url))
                    {
                        throw new ArgumentException("URL不能为空");
                    }

                    // 获取文件大小（如果尚未获取）
                    if (!totalBytes.HasValue)
                    {
                        using var headResponse = await GetHeadAsync(url, cancellationToken);
                        if (headResponse.IsSuccessStatusCode)
                        {
                            totalBytes = headResponse.Content.Headers.ContentLength;
                        }
                    }

                    // 确定总字节数 (赋值而非声明)
                    totalBytesValue = totalBytes ?? 0;

                    // 创建目录
                    string? directory = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // 创建HTTP请求
                    using var request = new HttpRequestMessage(HttpMethod.Get, url);

                    // 添加浏览器模拟头
                    request.Headers.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
                    request.Headers.Accept.ParseAdd("video/mp4, */*");
                    request.Headers.Referrer = new Uri("https://www.eeo.cn/");

                    if (startPosition > 0)  // 断点续传
                    {
                        request.Headers.Range = new RangeHeaderValue(startPosition, null);
                    }

                    // 发送请求
                    using var response = await _httpClient.SendAsync(
                        request,
                        HttpCompletionOption.ResponseHeadersRead,
                        cancellationToken);

                    // 检查响应状态
                    if (!response.IsSuccessStatusCode)
                    {
                        // 对于部分内容响应，即使状态码不是200也继续
                        if (response.StatusCode != HttpStatusCode.PartialContent)
                        {
                            throw new HttpRequestException($"服务器返回错误状态: {response.StatusCode}");
                        }
                    }

                    // 处理部分内容响应
                    if (response.StatusCode == HttpStatusCode.PartialContent)
                    {
                        var contentLength = response.Content.Headers.ContentLength;
                        if (contentLength.HasValue)
                        {
                            totalBytes = startPosition + contentLength.Value;
                            totalBytesValue = totalBytes.Value;
                        }
                    }
                    else if (!totalBytes.HasValue)  // 完整内容响应
                    {
                        totalBytes = response.Content.Headers.ContentLength;
                        totalBytesValue = totalBytes ?? 0;
                    }

                    // 读取内容流
                    using var contentStream = await response.Content.ReadAsStreamAsync();

                    // 设置缓冲区大小（基于线程数优化）
                    int bufferSize = Math.Max(65536, 1024 * 1024 / threads);
                    var buffer = new byte[bufferSize];
                    var lastBytesRead = startPosition;
                    var lastUpdateTime = DateTime.Now;

                    // 写入文件
                    using (var fileStream = new FileStream(
                        filePath,
                        startPosition > 0 ? FileMode.Append : FileMode.Create,
                        FileAccess.Write,
                        FileShare.None,
                        bufferSize,
                        true))
                    {
                        while (true)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            // 读取数据块
                            var bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length, cancellationToken);
                            if (bytesRead == 0) break;

                            // 写入文件
                            await fileStream.WriteAsync(buffer, 0, bytesRead, cancellationToken);

                            totalBytesRead += bytesRead;

                            // 优化进度更新（每200ms更新一次）
                            var elapsed = (DateTime.Now - lastUpdateTime).TotalMilliseconds;
                            if (elapsed > 200)
                            {
                                // 计算速度
                                var speed = (totalBytesRead - lastBytesRead) / (elapsed / 1000);
                                lastBytesRead = totalBytesRead;
                                lastUpdateTime = DateTime.Now;

                                // 计算剩余时间
                                var remainingTime = totalBytesValue > 0
                                    ? TimeSpan.FromSeconds((totalBytesValue - totalBytesRead) / (speed > 0 ? speed : 1))
                                    : TimeSpan.MaxValue;

                                // 报告进度
                                progressCallback?.Invoke(totalBytesRead, totalBytesValue, speed, remainingTime);
                            }
                        }

                        // 确保数据写入磁盘
                        await fileStream.FlushAsync();
                    }

                    // 确保最终进度报告
                    progressCallback?.Invoke(totalBytesValue, totalBytesValue, 0, TimeSpan.Zero);
                    isFinalUpdateSent = true;

                    // 添加延迟确保状态更新
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
                    // 服务器错误，可重试
                    retryCount++;
                    await Task.Delay(ExponentialBackoff(retryCount), cancellationToken);
                }
                catch (IOException ioEx) when (
                    ioEx is FileNotFoundException ||
                    ioEx is DirectoryNotFoundException ||
                    ioEx.HResult == -2147024894) // ERROR_FILE_NOT_FOUND
                {
                    // 文件路径问题，不可恢复
                    throw;
                }
                catch (Exception ex) when (retryCount < maxRetries)
                {
                    // 其他可恢复错误
                    Debug.WriteLine($"下载失败，正在重试({retryCount}/{maxRetries}): {ex.Message}");
                    retryCount++;
                    await Task.Delay(ExponentialBackoff(retryCount), cancellationToken);
                }
                finally
                {
                    // 确保发送最终进度更新
                    if (!isFinalUpdateSent)
                    {
                        progressCallback?.Invoke(totalBytesRead, totalBytesValue, 0, TimeSpan.Zero);
                        isFinalUpdateSent = true;
                    }
                }
            }

            return false;
        }

        // 指数退避算法（用于重试延迟） - 增加最大延迟
        private static int ExponentialBackoff(int retryCount)
        {
            return (int)Math.Min(1000 * Math.Pow(2, retryCount), 60000); // 最大60秒
        }

        // 取消指定URL的下载
        public static void CancelDownload(string url)
        {
            if (_activeDownloads.TryRemove(url, out var cts))
            {
                cts.Cancel();
            }
        }

        // 带并发控制的下载方法 - 确保信号量正确释放
        public static async Task<bool> DownloadWithConcurrencyControl(
            string url,
            string filePath,
            Action<long, long, double, TimeSpan> progressCallback,
            CancellationToken cancellationToken,
            int threads,
            Action? onDownloadStarted = null)
        {
            // 创建取消令牌
            var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            _activeDownloads[url] = cts;

            try
            {
                // 等待可用槽位
                await _concurrencySemaphore.WaitAsync(cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();

                // 调用开始回调
                onDownloadStarted?.Invoke();

                // 执行下载
                return await DownloadFileAsync(url, filePath, progressCallback, threads, cts.Token);
            }
            finally
            {
                // 确保释放信号量
                try
                {
                    _concurrencySemaphore.Release();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"释放信号量时出错: {ex.Message}");
                }

                // 从活动下载中移除
                _activeDownloads.TryRemove(url, out _);
                cts.Dispose();
            }
        }
    }

    // 原生方法封装（剪贴板操作）
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