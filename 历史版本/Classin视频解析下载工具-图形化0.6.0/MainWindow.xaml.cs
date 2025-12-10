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

namespace VideoDownloader
{
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
        private string downloadPath = string.Empty;
        private ObservableCollection<VideoItem> videoItems = new ObservableCollection<VideoItem>();
        private string lastClipboardText = string.Empty;
        private CancellationTokenSource? _clipboardCts;
        private bool _isClosing = false;
        private readonly int _maxConcurrentDownloads = 3;
        private readonly SemaphoreSlim _downloadSemaphore = new SemaphoreSlim(3);

        private double[] columnRatios = { 0.07, 0.3, 0.4, 0.25 };
        private bool userAdjustedColumns = false;
        private double lastTotalWidth = 0;

        private double[] originalColumnRatios = { 0.05, 0.3, 0.4, 0.25 };
        private GridLength originalPanelWidth;
        private GridLength originalResultWidth;
        private GridLength originalMainContentHeight;
        private GridLength originalLogHeight;

        private bool _logScrollToEnd = true;
        private bool _outputScrollToEnd = true;

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

        public class VideoItem : INotifyPropertyChanged
        {
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
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            private string _downloadedSize = "0 B";
            public string DownloadedSize
            {
                get => _downloadedSize;
                set
                {
                    if (_downloadedSize != value)
                    {
                        _downloadedSize = value;
                        OnPropertyChanged(nameof(DownloadedSize));
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            private string _totalSize = "0 B";
            public string TotalSize
            {
                get => _totalSize;
                set
                {
                    if (_totalSize != value)
                    {
                        _totalSize = value;
                        OnPropertyChanged(nameof(TotalSize));
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            private string _currentSpeed = "0 MB/s";
            public string CurrentSpeed
            {
                get => _currentSpeed;
                set
                {
                    if (_currentSpeed != value)
                    {
                        _currentSpeed = value;
                        OnPropertyChanged(nameof(CurrentSpeed));
                        OnPropertyChanged(nameof(DisplayStatus));
                    }
                }
            }

            private string _remainingTime = "未知";
            public string RemainingTime
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

            public string DisplayStatus
            {
                get => Status == "下载中..." ?
                    $"下载中 ({Progress}%) - {DownloadedSize}/{TotalSize} @ {CurrentSpeed} - 剩余: {RemainingTime}" :
                    Status;
            }

            public Brush StatusColor { get; set; } = Brushes.Gray;
            public bool IsDownloading { get; set; } = false;
            public bool IsDownloaded { get; set; } = false;
            public CancellationTokenSource? DownloadTokenSource { get; set; }

            public event PropertyChangedEventHandler? PropertyChanged;

            protected virtual void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            InitializeApp();
            this.Closed += MainWindow_Closed;

            originalPanelWidth = PanelColumn.Width;
            originalResultWidth = ResultColumn.Width;
            originalMainContentHeight = MainContentRow.Height;
            originalLogHeight = LogRow.Height;
            Array.Copy(columnRatios, originalColumnRatios, columnRatios.Length);

            DisableCloseButton();
        }

        private void InitializeTool_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (videoItems.Any(item => item.IsDownloading))
                {
                    if (MessageBox.Show("当前有视频正在下载中，继续初始化将取消所有下载任务。\n是否继续初始化？",
                        "警告", MessageBoxButton.OKCancel, MessageBoxImage.Warning) != MessageBoxResult.OK)
                    {
                        AppendLog("用户取消了初始化操作");
                        return;
                    }
                    CancelAllDownloads();
                }

                AppendLog("开始初始化解析工具...");

                txtLog.Clear();
                txtDownloadOutput.Clear();
                CancelAllDownloads();
                videoItems.Clear();

                Dispatcher.Invoke(() =>
                {
                    lstResults.ItemsSource = null;
                    lstResults.ItemsSource = videoItems;
                    UpdateVideoCountDisplay();
                });

                this.Width = 1200;
                this.Height = 675;
                progressBar.Value = 0;
                txtStatus.Text = "准备就绪...";
                txtStatus.Foreground = Brushes.Black;
                txtCurrentAction.Text = "就绪";
                ResetLayoutSizes();
                GenerateBatFile();

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

        private T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T result) return result;

                var descendant = FindVisualChild<T>(child);
                if (descendant != null) return descendant;
            }
            return null;
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
                AppendLog($"禁用关闭按钮时出错: {ex.Message}");
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
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            logScrollViewer.ScrollChanged += LogScrollViewer_ScrollChanged;
            downloadOutputScrollViewer.ScrollChanged += OutputScrollViewer_ScrollChanged;
            txtStatus.Text = "应用程序已就绪";
            txtCurrentAction.Text = "请开始操作";
        }

        private void LogScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            _logScrollToEnd = Math.Abs(e.VerticalOffset - (e.ExtentHeight - e.ViewportHeight)) < 1;
        }

        private void OutputScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            _outputScrollToEnd = Math.Abs(e.VerticalOffset - (e.ExtentHeight - e.ViewportHeight)) < 1;
        }

        private void MainWindow_Closed(object? sender, EventArgs e)
        {
            _clipboardCts?.Cancel();
            CancelAllDownloads();
        }

        private void CancelAllDownloads()
        {
            foreach (var item in videoItems)
            {
                if (item.IsDownloading && item.DownloadTokenSource != null)
                {
                    item.DownloadTokenSource.Cancel();
                }
            }
        }

        private void InitializeApp()
        {
            downloadPath = Path.Combine(Environment.CurrentDirectory, "下载目录");
            Directory.CreateDirectory(downloadPath);
            UpdateDownloadPathDisplay();
            File.WriteAllText("解析结果.ini", string.Empty);
            GenerateBatFile();
            UpdateVideoCountDisplay();
            AppendLog("应用程序已初始化");
        }

        private void UpdateVideoCountDisplay()
        {
            Dispatcher.Invoke(() => txtTotalVideoCount.Text = $"共解析了 {videoItems.Count} 个视频");
        }

        private void UpdateDownloadPathDisplay()
        {
            txtDownloadPath.Text = downloadPath;
            txtDownloadPath.ToolTip = downloadPath;
        }

        private void GenerateBatFile()
        {
            try
            {
                string batContent = "@echo off\r\nchcp 65001\r\n";

                foreach (var item in videoItems)
                {
                    if (!string.IsNullOrEmpty(item.Url))
                    {
                        string safeName = CleanName(item.Name);
                        batContent += $"aria2c.exe --allow-overwrite=true -d \"{downloadPath}\" -o \"{safeName}.mp4\" \"{item.Url}\"\r\n";
                    }
                }

                File.WriteAllText("下载视频工具.bat", batContent);
            }
            catch (Exception ex)
            {
                UpdateStatus($"更新下载脚本失败: {ex.Message}", Brushes.Red);
            }
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
                    Dispatcher.Invoke(() =>
                    {
                        txtCurrentAction.Text = "已将 'getLessonRecordInfo' 复制到剪贴板";
                        txtStatus.Text = "请到浏览器开发者工具中执行此命令";
                        txtStatus.Foreground = Brushes.Green;
                    });
                    progressBar.Value = 100;
                    AppendLog("已成功将命令复制到剪贴板");
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

        private void UpdateStatus(string message, Brush color)
        {
            Dispatcher.Invoke(() =>
            {
                txtStatus.Text = message;
                txtStatus.Foreground = color;
            });
        }

        private void Parse_Click(object sender, RoutedEventArgs e)
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

                lastClipboardText = Clipboard.GetText();
                ParseContent(lastClipboardText);
            }
            catch (Exception ex)
            {
                UpdateStatus($"解析失败: {ex.Message}", Brushes.Red);
            }
        }

        private void ParseContent(string content)
        {
            List<string> duplicateFiles = new List<string>();
            List<string> duplicateCourses = new List<string>();

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
                        return;
                    }

                    if (videoItems.Any(item => string.Equals(item.Name, lessonName, StringComparison.OrdinalIgnoreCase)))
                    {
                        duplicateCourses.Add(lessonName);
                        return;
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

                    var newItem = new VideoItem
                    {
                        Name = lessonName,
                        Status = "解析完成",
                        StatusColor = Brushes.Green
                    };

                    if (!string.IsNullOrEmpty(lastValidUrl))
                    {
                        newItem.Url = lastValidUrl;
                    }
                    else
                    {
                        newItem.Status = "缺少MP4 URL";
                        newItem.StatusColor = Brushes.Orange;
                    }

                    videoItems.Add(newItem);
                    UpdateVideoIndexes();

                    Dispatcher.Invoke(() =>
                    {
                        lstResults.ItemsSource = null;
                        lstResults.ItemsSource = videoItems;
                        UpdateVideoCountDisplay();
                    });

                    if (!string.IsNullOrEmpty(lastValidUrl))
                    {
                        UpdateStatus($"成功解析视频", Brushes.Green);
                        GenerateBatFile();
                    }
                    else
                    {
                        UpdateStatus("未找到视频信息", Brushes.Orange);
                    }
                }
                else
                {
                    UseOriginalLineParsing(content, duplicateFiles, duplicateCourses);
                }

                progressBar.Value = 100;
                txtCurrentAction.Text = "解析完成";
            }
            catch (JsonException)
            {
                UseOriginalLineParsing(content, duplicateFiles, duplicateCourses);
            }
            finally
            {
                ShowDuplicateMessage(duplicateFiles, duplicateCourses);
            }
        }

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
                if (duplicateCourses.Count > 5) message.AppendLine($"...及其他 {duplicateCourses.Count - 5} 个");
            }

            if (message.Length > 0)
            {
                Dispatcher.Invoke(() => MessageBox.Show(message.ToString(), "重复内容提示",
                    MessageBoxButton.OK, MessageBoxImage.Information));
            }
        }

        private void UseOriginalLineParsing(string content, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            string[] lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            bool foundLessonName = false;
            string currentLessonName = string.Empty;
            string finalUrl = string.Empty;
            bool playsetEncountered = false;
            bool inFileItem = false;

            foreach (string line in lines)
            {
                if (line.Contains("lessonName", StringComparison.OrdinalIgnoreCase))
                {
                    currentLessonName = ExtractValue(line, "lessonName");
                    foundLessonName = true;
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
                    return;
                }

                if (videoItems.Any(item => string.Equals(item.Name, currentLessonName, StringComparison.OrdinalIgnoreCase)))
                {
                    duplicateCourses.Add(currentLessonName);
                    return;
                }

                VideoItem newItem = new VideoItem
                {
                    Name = currentLessonName,
                    Status = "解析完成",
                    StatusColor = Brushes.Green
                };

                if (!string.IsNullOrEmpty(finalUrl))
                {
                    newItem.Url = finalUrl;
                }
                else
                {
                    newItem.Status = "缺少MP4 URL";
                    newItem.StatusColor = Brushes.Orange;
                }

                videoItems.Add(newItem);
                UpdateVideoIndexes();

                Dispatcher.Invoke(() =>
                {
                    lstResults.ItemsSource = null;
                    lstResults.ItemsSource = videoItems;
                    UpdateVideoCountDisplay();
                });
            }
            else
            {
                UpdateStatus("未找到课程名称", Brushes.Orange);
            }

            progressBar.Value = 100;
            txtCurrentAction.Text = "解析完成";
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

        private void Download_Click(object sender, RoutedEventArgs e)
        {
            if (videoItems.Count == 0)
            {
                UpdateStatus("没有可下载的视频", Brushes.Orange);
                return;
            }

            try
            {
                txtDownloadOutput.Clear();
                var tabControl = FindParentTabControl(txtDownloadOutput);
                if (tabControl != null && tabControl.Items.Count > 1) tabControl.SelectedIndex = 1;

                txtCurrentAction.Text = "正在启动下载...";
                progressBar.Value = 50;
                AppendDownloadOutput("开始下载视频...");
                AppendDownloadOutput($"最大同时下载数: {_maxConcurrentDownloads}");

                if (!File.Exists("aria2c.exe"))
                {
                    UpdateStatus("错误: 未找到aria2c.exe", Brushes.Red);
                    return;
                }

                Task.Run(() => DownloadVideosAsync());
            }
            catch (Exception ex)
            {
                UpdateStatus($"启动下载失败: {ex.Message}", Brushes.Red);
            }
        }

        private async void DownloadSingle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is VideoItem item)
            {
                try
                {
                    txtDownloadOutput.Clear();
                    var tabControl = FindParentTabControl(txtDownloadOutput);
                    if (tabControl != null && tabControl.Items.Count > 1) tabControl.SelectedIndex = 1;

                    txtCurrentAction.Text = $"正在下载: {Truncate(item.Name, 20)}";
                    AppendLog($"开始下载单个视频: {item.Name}");

                    // 添加状态更新
                    Dispatcher.Invoke(() =>
                    {
                        item.Status = "下载中...";
                        item.Progress = 0;
                        item.DownloadedSize = "0 B";
                        item.TotalSize = "0 B";
                        item.CurrentSpeed = "0 MB/s";
                        item.RemainingTime = "未知";
                        item.StatusColor = Brushes.Orange;
                        item.IsDownloading = true;
                        lstResults.Items.Refresh();
                    });

                    if (!File.Exists("aria2c.exe"))
                    {
                        UpdateStatus("错误: 未找到aria2c.exe", Brushes.Red);
                        return;
                    }

                    await DownloadSingleVideoAsync(item, new CancellationTokenSource().Token);
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
                Dispatcher.Invoke(() => txtCurrentAction.Text = "下载进行中...");

                var videosToDownload = videoItems
                    .Where(item => !item.IsDownloading && !item.IsDownloaded)
                    .ToList();

                if (videosToDownload.Count == 0)
                {
                    AppendDownloadOutput("没有需要下载的视频");
                    Dispatcher.Invoke(() => txtCurrentAction.Text = "没有需要下载的任务");
                    return;
                }

                int successCount = 0;
                int failedCount = 0;
                int skippedCount = 0;
                var downloadTasks = new List<Task>();

                foreach (var item in videosToDownload)
                {
                    string safeName = CleanName(item.Name);
                    string outputFile = Path.Combine(downloadPath, $"{safeName}.mp4");

                    if (File.Exists(outputFile))
                    {
                        duplicateFiles.Add($"{safeName}.mp4");
                        skippedCount++;
                        Dispatcher.Invoke(() =>
                        {
                            item.Status = "下载完成";
                            item.Progress = 100;
                            item.StatusColor = Brushes.Green;
                            item.IsDownloaded = true;
                            lstResults.Items.Refresh();
                        });
                        continue;
                    }

                    item.DownloadTokenSource = new CancellationTokenSource();
                    var token = item.DownloadTokenSource.Token;

                    var task = Task.Run(async () =>
                    {
                        try
                        {
                            await _downloadSemaphore.WaitAsync(token);

                            Dispatcher.Invoke(() =>
                            {
                                item.Status = "下载中...";
                                item.Progress = 0;
                                item.DownloadedSize = "0 B";
                                item.TotalSize = "0 B";
                                item.CurrentSpeed = "0 MB/s";
                                item.RemainingTime = "未知";
                                item.StatusColor = Brushes.Orange;
                                item.IsDownloading = true;
                                lstResults.Items.Refresh();
                            });

                            bool result = await DownloadSingleVideoAsync(item, token);

                            if (result) Interlocked.Increment(ref successCount);
                            else Interlocked.Increment(ref failedCount);
                        }
                        catch (OperationCanceledException)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                item.Status = "等待下载";
                                item.StatusColor = Brushes.Orange;
                                item.IsDownloading = false;
                                lstResults.Items.Refresh();
                            });
                        }
                        finally
                        {
                            _downloadSemaphore.Release();
                        }
                    }, token);

                    downloadTasks.Add(task);
                }

                await Task.WhenAll(downloadTasks);
                AppendDownloadOutput($"下载完成: 成功 {successCount} 个, 失败 {failedCount} 个, 跳过 {skippedCount} 个");

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
                AppendDownloadOutput($"下载过程中出错: {ex.Message}");
                Dispatcher.Invoke(() => UpdateStatus($"下载失败: {ex.Message}", Brushes.Red));
            }
        }

        private void ShowDuplicateFilesMessage(List<string> duplicateFiles)
        {
            if (duplicateFiles.Count == 0) return;

            StringBuilder message = new StringBuilder();
            message.AppendLine($"跳过 {duplicateFiles.Count} 个已存在的视频文件：");

            foreach (var file in duplicateFiles.Take(5))
            {
                message.AppendLine($"- {Truncate(file, 50)}");
            }

            if (duplicateFiles.Count > 5)
            {
                message.AppendLine($"...及其他 {duplicateFiles.Count - 5} 个");
            }

            Dispatcher.Invoke(() =>
            {
                MessageBox.Show(message.ToString(),
                    "文件已存在",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            });
        }

        private async Task<bool> DownloadSingleVideoAsync(VideoItem item, CancellationToken token)
        {
            try
            {
                if (string.IsNullOrEmpty(item.Url))
                {
                    Dispatcher.Invoke(() => item.Status = "下载失败: 无URL");
                    return false;
                }

                string safeName = CleanName(item.Name);
                string outputFile = Path.Combine(downloadPath, $"{safeName}.mp4");

                if (File.Exists(outputFile))
                {
                    Dispatcher.Invoke(() =>
                    {
                        item.Status = "下载完成";
                        item.Progress = 100;
                        item.StatusColor = Brushes.Green;
                        item.IsDownloaded = true;
                        lstResults.Items.Refresh();
                    });
                    return true;
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = "aria2c.exe",
                    Arguments = $"--allow-overwrite=true -d \"{downloadPath}\" -o \"{safeName}.mp4\" \"{item.Url}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                using (var process = new Process { StartInfo = startInfo })
                {
                    process.OutputDataReceived += (s, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            AppendDownloadOutput(e.Data);
                            ParseDownloadProgress(e.Data, item);
                        }
                    };

                    process.ErrorDataReceived += (s, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data)) AppendDownloadOutput($"错误: {e.Data}");
                    };

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();
                    await process.WaitForExitAsync(token);

                    if (process.ExitCode == 0)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            item.Status = "下载完成";
                            item.Progress = 100;
                            item.StatusColor = Brushes.Green;
                            item.IsDownloaded = true;
                            lstResults.Items.Refresh();
                        });
                        return true;
                    }
                    else
                    {
                        Dispatcher.Invoke(() =>
                        {
                            item.Status = "下载失败";
                            item.StatusColor = Brushes.Red;
                            lstResults.Items.Refresh();
                        });
                        return false;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                Dispatcher.Invoke(() => item.Status = "已取消");
                return false;
            }
            catch (Exception ex)
            {
                AppendLog($"下载失败: {ex.Message}");
                Dispatcher.Invoke(() =>
                {
                    item.Status = "下载失败";
                    item.StatusColor = Brushes.Red;
                    lstResults.Items.Refresh();
                });
                return false;
            }
        }

        private void ParseDownloadProgress(string data, VideoItem item)
        {
            try
            {
                string pattern = @"\[\#\w+\s+([\d.]+\w*B)/([\d.]+\w*B)\((\d+)%\)\s+CN:\d+\s+DL:([\d.]+\w*B)\s+ETA:([\d]+\w*)\]";
                var match = Regex.Match(data, pattern);

                if (match.Success && match.Groups.Count == 6)
                {
                    string downloadedSize = ConvertToDecimalSize(match.Groups[1].Value);
                    string totalSize = ConvertToDecimalSize(match.Groups[2].Value);
                    int progress = int.Parse(match.Groups[3].Value);
                    string currentSpeed = ConvertToMBps(match.Groups[4].Value);
                    string remainingTime = ConvertTimeUnitsToChinese(match.Groups[5].Value);

                    Dispatcher.Invoke(() =>
                    {
                        item.DownloadedSize = downloadedSize;
                        item.TotalSize = totalSize;
                        item.Progress = progress;
                        item.CurrentSpeed = currentSpeed;
                        item.RemainingTime = remainingTime;
                    });
                }
            }
            catch { }
        }

        private string ConvertTimeUnitsToChinese(string timeString)
        {
            if (timeString.Contains("m") && timeString.Contains("s"))
            {
                var parts = timeString.Split('m', 's');
                if (parts.Length >= 2 && int.TryParse(parts[0], out int minutes) && int.TryParse(parts[1], out int seconds))
                    return $"{minutes}分{seconds}秒";
            }
            else if (timeString.Contains("m"))
            {
                var parts = timeString.Split('m');
                if (int.TryParse(parts[0], out int minutes)) return $"{minutes}分";
            }
            else if (timeString.Contains("s"))
            {
                var parts = timeString.Split('s');
                if (int.TryParse(parts[0], out int seconds)) return $"{seconds}秒";
            }
            else if (timeString.Contains("h"))
            {
                var parts = timeString.Split('h');
                if (int.TryParse(parts[0], out int hours)) return $"{hours}小时";
            }
            return timeString;
        }

        private string ConvertToMBps(string speedWithUnit)
        {
            if (string.IsNullOrEmpty(speedWithUnit)) return "0 MB/s";

            Match match = Regex.Match(speedWithUnit, @"([\d.]+)(\w*)");
            if (!match.Success || match.Groups.Count < 3) return "0 MB/s";

            double value = double.Parse(match.Groups[1].Value);
            string unit = match.Groups[2].Value.ToUpper();
            double valueInMBps = 0;

            switch (unit)
            {
                case "B": valueInMBps = value / 1000000.0; break;
                case "KIB": valueInMBps = value * 1024 / 1000000.0; break;
                case "MIB": valueInMBps = value * 1048576 / 1000000.0; break;
                case "GIB": valueInMBps = value * 1073741824 / 1000000000.0; break;
                default: return speedWithUnit;
            }

            return $"{valueInMBps:0.00} MB/s";
        }

        private string ConvertToDecimalSize(string sizeWithUnit)
        {
            if (string.IsNullOrEmpty(sizeWithUnit)) return "0 B";

            Match match = Regex.Match(sizeWithUnit, @"([\d.]+)(\w*)");
            if (!match.Success || match.Groups.Count < 3) return sizeWithUnit;

            double value = double.Parse(match.Groups[1].Value);
            string unit = match.Groups[2].Value.ToUpper();
            double bytes = 0;

            switch (unit)
            {
                case "B": bytes = value; break;
                case "KIB": bytes = value * 1024; break;
                case "MIB": bytes = value * 1048576; break;
                case "GIB": bytes = value * 1073741824; break;
                default: return sizeWithUnit;
            }

            if (bytes >= 1000000000) return $"{bytes / 1000000000:0.00} GB";
            if (bytes >= 1000000) return $"{bytes / 1000000:0.00} MB";
            if (bytes >= 1000) return $"{bytes / 1000:0.00} KB";
            return $"{bytes:0} B";
        }

        private TabControl? FindParentTabControl(DependencyObject child)
        {
            DependencyObject parent = VisualTreeHelper.GetParent(child);
            while (parent != null && !(parent is TabControl))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }
            return parent as TabControl;
        }

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
                    GenerateBatFile();
                    UpdateStatus($"下载路径已更新: {downloadPath}", Brushes.Green);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"更改目录失败: {ex.Message}", Brushes.Red);
            }
        }

        private async void CopyRawData_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(lastClipboardText))
            {
                UpdateStatus("没有可复制的原始数据", Brushes.Orange);
                return;
            }

            try
            {
                bool success = await SafeSetClipboard(lastClipboardText, CancellationToken.None);
                if (success) UpdateStatus("原始数据已复制到剪贴板", Brushes.Green);
                else UpdateStatus("复制原始数据失败", Brushes.Red);
            }
            catch (Exception ex)
            {
                UpdateStatus($"复制原始数据失败: {ex.Message}", Brushes.Red);
            }
        }

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

        private void DeleteItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is VideoItem item)
            {
                try
                {
                    if (item.IsDownloading)
                    {
                        if (MessageBox.Show($"视频 '{Truncate(item.Name, 50)}' 正在下载中，删除将取消下载。\n是否继续删除？",
                            "警告", MessageBoxButton.OKCancel, MessageBoxImage.Warning) != MessageBoxResult.OK)
                        {
                            return;
                        }
                        item.DownloadTokenSource?.Cancel();
                    }

                    videoItems.Remove(item);
                    UpdateVideoIndexes();
                    Dispatcher.Invoke(() =>
                    {
                        lstResults.Items.Refresh();
                        UpdateVideoCountDisplay();
                    });
                    GenerateBatFile();
                }
                catch (Exception ex)
                {
                    AppendLog($"删除视频项失败: {ex.Message}");
                }
            }
        }

        private void AppendLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                message = message.Replace("\r", "").Replace("\n", " ");
                string logEntry = $"{DateTime.Now:HH:mm:ss} - {message}";

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
            });
        }

        private void AppendDownloadOutput(string message)
        {
            Dispatcher.Invoke(() =>
            {
                message = message.Replace("\r", "").Replace("\n", " ");

                if (!string.IsNullOrEmpty(txtDownloadOutput.Text) &&
                   !txtDownloadOutput.Text.EndsWith(Environment.NewLine))
                {
                    txtDownloadOutput.AppendText(Environment.NewLine);
                }
                txtDownloadOutput.AppendText(message);

                if (_outputScrollToEnd)
                {
                    txtDownloadOutput.CaretIndex = txtDownloadOutput.Text.Length;
                    downloadOutputScrollViewer.ScrollToEnd();
                }
            });
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

    public static class ProcessExtensions
    {
        public static Task WaitForExitAsync(this Process process, CancellationToken cancellationToken = default)
        {
            var tcs = new TaskCompletionSource<object?>();
            process.EnableRaisingEvents = true;
            process.Exited += (sender, args) => tcs.TrySetResult(null);

            if (cancellationToken != default)
            {
                cancellationToken.Register(() => tcs.TrySetCanceled());
            }

            return tcs.Task;
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