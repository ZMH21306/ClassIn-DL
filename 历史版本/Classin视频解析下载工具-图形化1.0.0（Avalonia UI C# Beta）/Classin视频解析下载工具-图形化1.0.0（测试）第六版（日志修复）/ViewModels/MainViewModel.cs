using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.Input;
using Classin视频解析下载工具.Commands;
using Classin视频解析下载工具.Constants;
using Classin视频解析下载工具.Models;
using Classin视频解析下载工具.Services;
using Classin视频解析下载工具.ViewModels;

namespace Classin视频解析下载工具.ViewModels
{
    public partial class MainViewModel : ViewModelBase, IDisposable
    {
        private readonly IClipboardService _clipboardService;
        private readonly IValidationService _validationService;
        private readonly IDuplicateDetectionService _duplicateDetectionService;
        private readonly IConfigurationManager _configurationManager;
        private readonly IHealthMonitor _healthMonitor;
        private readonly IDownloadService _downloadService;
        private readonly IUIService _uiService;
        private readonly IFileFolderService _fileFolderService;
        private readonly IDialogService _dialogService;
        private readonly IParseService _parseService;
        private readonly ILoggingService _loggingService;
        private readonly LogViewModel _logViewModel;
        private readonly DownloadManager _downloadManager;
        
        private string _downloadPath = string.Empty;
        private int _maxConcurrentDownloads = 3; // 默认值改为3，与WPF版本一致
        private string _statusSummary = "当前没有视频正在下载，统计信息为空";
        private bool _isInitialized;
        private bool _disposed;

        public MainViewModel()
        {
            // 默认构造函数用于设计时支持
            _clipboardService = new ClipboardService();
            _validationService = new ValidationService();
            _duplicateDetectionService = new DuplicateDetectionService();
            _configurationManager = new ConfigurationManager();
            _healthMonitor = new HealthMonitor();
            _downloadService = new DownloadService(_validationService);
            _uiService = new UIService();
            _fileFolderService = new FileFolderService();
            _dialogService = new DialogService();
            _parseService = new ParseService(_validationService, _duplicateDetectionService, _downloadService);
            _loggingService = new LoggingService();
            _logViewModel = new LogViewModel(_loggingService);
            _downloadManager = new DownloadManager(_downloadService, _duplicateDetectionService, new ProgressTrackingService(), _loggingService);
            
            InitializeCommands();
            LoadConfiguration();
            SubscribeToDownloadManager();
        }

        public MainViewModel(
            IClipboardService clipboardService,
            IValidationService validationService,
            IDuplicateDetectionService duplicateDetectionService,
            IConfigurationManager configurationManager,
            IHealthMonitor healthMonitor,
            IDownloadService downloadService,
            IUIService uiService,
            IFileFolderService fileFolderService,
            IDialogService dialogService,
            IParseService parseService,
            ILoggingService loggingService,
            LogViewModel logViewModel,
            DownloadManager downloadManager)
        {
            _clipboardService = clipboardService;
            _validationService = validationService;
            _duplicateDetectionService = duplicateDetectionService;
            _configurationManager = configurationManager;
            _healthMonitor = healthMonitor;
            _downloadService = downloadService;
            _uiService = uiService;
            _fileFolderService = fileFolderService;
            _dialogService = dialogService;
            _parseService = parseService;
            _loggingService = loggingService;
            _logViewModel = logViewModel;
            _downloadManager = downloadManager;
            
            InitializeCommands();
            LoadConfiguration();
            SubscribeToDownloadManager();
        }

        public ObservableCollection<VideoItem> VideoItems => _downloadManager.VideoItems;

        public string DownloadPath
        {
            get => _downloadPath;
            set
            {
                if (SetProperty(ref _downloadPath, value))
                {
                    _downloadManager.UpdatePath(value);
                    _configurationManager.UpdateConfiguration(c => c.DownloadPath = value);
                }
            }
        }

        public int MaxConcurrentDownloads
        {
            get => _maxConcurrentDownloads;
            set
            {
                // 确保值在有效范围内
                int clampedValue = value;
                if (clampedValue < 1) clampedValue = 1;
                if (clampedValue > 16) clampedValue = 16;
                
                if (SetProperty(ref _maxConcurrentDownloads, clampedValue))
                {
                    _downloadManager.MaxConcurrentDownloads = clampedValue;
                    _configurationManager.UpdateConfiguration(c => c.MaxConcurrentDownloads = clampedValue);
                    _loggingService.Info($"配置已更新: 最大并发下载数 = {clampedValue}", "Configuration");
                    OnPropertyChanged(nameof(MaxConcurrentDownloadsText)); // 强制更新显示文本
                }
            }
        }

        public string MaxConcurrentDownloadsText => $"当前值: {_maxConcurrentDownloads}";

        public string StatusSummary
        {
            get => _statusSummary;
            set => SetProperty(ref _statusSummary, value);
        }

        public string LogText => _logViewModel.FilteredLogText;

        public bool IsInitialized
        {
            get => _isInitialized;
            set => SetProperty(ref _isInitialized, value);
        }

        // Commands
        public Classin视频解析下载工具.Commands.RelayCommand CopyCommandCommand { get; private set; } = null!;
        public Classin视频解析下载工具.Commands.RelayCommand ParseCommandCommand { get; private set; } = null!;
        public Classin视频解析下载工具.Commands.RelayCommand DownloadAllCommand { get; private set; } = null!;
        public Classin视频解析下载工具.Commands.RelayCommand OpenFolderCommand { get; private set; } = null!;
        public Classin视频解析下载工具.Commands.RelayCommand ChangeDownloadPathCommand { get; private set; } = null!;
        public Classin视频解析下载工具.Commands.RelayCommand InitializeToolCommand { get; private set; } = null!;
        public Classin视频解析下载工具.Commands.RelayCommand<VideoItem> CopyUrlCommand { get; private set; } = null!;
        public Classin视频解析下载工具.Commands.RelayCommand<VideoItem> DownloadSingleCommand { get; private set; } = null!;
        public Classin视频解析下载工具.Commands.RelayCommand<VideoItem> DeleteItemCommand { get; private set; } = null!;
        public Classin视频解析下载工具.Commands.RelayCommand<double> SliderValueChangedCommand { get; private set; } = null!;

        private void InitializeCommands()
        {
            CopyCommandCommand = new Classin视频解析下载工具.Commands.RelayCommand(async () => await ExecuteCopyCommandAsync());
            ParseCommandCommand = new Classin视频解析下载工具.Commands.RelayCommand(async () => await ExecuteParseCommandAsync());
            DownloadAllCommand = new Classin视频解析下载工具.Commands.RelayCommand(async () => await ExecuteDownloadAllCommandAsync());
            OpenFolderCommand = new Classin视频解析下载工具.Commands.RelayCommand(ExecuteOpenFolderCommand);
            ChangeDownloadPathCommand = new Classin视频解析下载工具.Commands.RelayCommand(async () => await ExecuteChangeDownloadPathCommandAsync());
            InitializeToolCommand = new Classin视频解析下载工具.Commands.RelayCommand(async () => await ExecuteInitializeToolCommandAsync());
            CopyUrlCommand = new Classin视频解析下载工具.Commands.RelayCommand<VideoItem>(ExecuteCopyUrlCommand);
            DownloadSingleCommand = new Classin视频解析下载工具.Commands.RelayCommand<VideoItem>(ExecuteDownloadSingleCommand);
            DeleteItemCommand = new Classin视频解析下载工具.Commands.RelayCommand<VideoItem>(ExecuteDeleteItemCommand);
            SliderValueChangedCommand = new Classin视频解析下载工具.Commands.RelayCommand<double>(ExecuteSliderValueChangedCommand);
        }

        private void LoadConfiguration()
        {
            var config = _configurationManager.GetConfiguration();
            MaxConcurrentDownloads = config.MaxConcurrentDownloads;
            DownloadPath = config.DownloadPath;
        }

        private void SubscribeToDownloadManager()
        {
            _downloadManager.PropertyChanged += OnDownloadManagerPropertyChanged;
        }

        private void OnDownloadManagerPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case nameof(DownloadManager.StatusSummary):
                    StatusSummary = _downloadManager.StatusSummary;
                    break;
            }
        }

        public void Initialize(string baseDirectory)
        {
            _healthMonitor.RecordInitialization();

            // 只有当下载路径为空时，才设置默认路径
            if (string.IsNullOrEmpty(DownloadPath))
            {
                var newDownloadPath = Path.Combine(baseDirectory, AppConstants.DefaultDownloadFolder);
                DownloadPath = newDownloadPath;
            }
            
            // 确保下载目录存在，添加错误处理
            try
            {
                Directory.CreateDirectory(DownloadPath);
                // 先设置初始化标志，确保日志可以显示
                _logViewModel.IsInitialized = true;
                AppendLog($"下载目录已创建: {DownloadPath}");
            }
            catch (Exception ex)
            {
                _logViewModel.IsInitialized = true;  // 发生错误时也要设置
                AppendLog($"创建下载目录失败: {ex.Message}");
                // 即使目录创建失败，也继续初始化，避免程序启动失败
            }

            _downloadManager.Initialize(DownloadPath);
                        
            IsInitialized = true;
            AppendLog("软件初始化完成");
        }

        private async Task ExecuteCopyCommandAsync()
        {
            try
            {
                _loggingService.Info("用户点击复制筛选关键词按钮", "UI.Command");
                AppendLog("点击按钮：复制筛选关键词");
                bool success = await _clipboardService.SetClipboardTextAsync("getLessonRecordInfo");
                if (success)
                {
                    _loggingService.Info("筛选关键词已成功复制到剪贴板", "Clipboard");
                    AppendLog("筛选文本已成功复制到剪贴板");
                }
                else
                {
                    _loggingService.Warning("筛选关键词复制失败，建议用户手动复制", "Clipboard");
                    AppendLog("复制失败，请手动复制命令");
                }
            }
            catch (Exception ex)
            {
                _loggingService.Error($"复制筛选关键词时发生异常: {ex.Message}", "Clipboard");
                AppendLog($"复制命令失败: {ex.Message}");
            }
        }

        private async Task ExecuteParseCommandAsync()
        {
            try
            {
                _loggingService.Info("用户点击解析请求头按钮", "UI.Command");
                AppendLog("点击按钮：解析请求头");
                if (!_clipboardService.ContainsText())
                {
                    _loggingService.Warning("剪贴板中没有文本内容，无法进行解析", "Parse");
                    AppendLog("剪贴板中没有文本内容");
                    return;
                }

                string clipboardText = await _clipboardService.GetTextAsync() ?? string.Empty;
                if (string.IsNullOrEmpty(clipboardText))
                {
                    _loggingService.Warning("无法从剪贴板获取内容", "Parse");
                    AppendLog("无法获取剪贴板内容");
                    return;
                }

                _loggingService.Debug($"开始解析剪贴板内容，长度: {clipboardText.Length} 字符", "Parse");
                
                // 简化的解析逻辑
                var parseSuccess = await ParseContentAsync(clipboardText);

                if (parseSuccess)
                {
                    _loggingService.Info("请求头解析成功完成", "Parse");
                    AppendLog("请求头解析成功");
                }
                else
                {
                    _loggingService.Warning("请求头解析失败", "Parse");
                    AppendLog("请求头解析失败");
                }
            }
            catch (Exception ex)
            {
                _loggingService.Error($"解析请求头时发生异常: {ex.Message}", "Parse");
                AppendLog($"解析失败: {ex.Message}");
            }
        }

        private async Task<bool> ParseContentAsync(string content)
        {
            var duplicateFiles = new List<string>();
            var duplicateCourses = new List<string>();

            try
            {
                return await TryParseJsonContentAsync(content, duplicateFiles, duplicateCourses);
            }
            catch
            {
                return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
            }
            finally
            {
                ShowDuplicateMessage(duplicateFiles, duplicateCourses);
            }
        }

        private async Task<bool> TryParseJsonContentAsync(
            string content, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(content);
                var root = doc.RootElement;

                if (!root.TryGetProperty("data", out var data))
                {
                    return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
                }

                var lessonName = data.TryGetProperty("lessonName", out var lessonNameElement)
                    ? lessonNameElement.GetString() ?? string.Empty
                    : string.Empty;

                if (string.IsNullOrEmpty(lessonName))
                {
                    return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
                }

                var (isDuplicate, _) = await CheckForDuplicatesAsync(lessonName, duplicateFiles, duplicateCourses);
                if (isDuplicate)
                {
                    return false;
                }

                var lastValidUrl = ExtractVideoUrlFromJson(data);
                long fileSize = await GetFileSizeIfNeededAsync(lastValidUrl);

                return await CreateAndAddVideoItemAsync(lessonName, lastValidUrl, fileSize);
            }
            catch
            {
                return await UseOriginalLineParsingAsync(content, duplicateFiles, duplicateCourses);
            }
        }

        private async Task<(bool isDuplicate, bool duplicateFound)> CheckForDuplicatesAsync(
            string lessonName, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            var safeName = _duplicateDetectionService.SanitizeFileName(lessonName);
            var outputFile = Path.Combine(DownloadPath, $"{safeName}.mp4");

            if (File.Exists(outputFile))
            {
                duplicateFiles.Add($"{safeName}.mp4");
                AppendLog($"已跳过重复视频项: {lessonName}");
                return (true, true);
            }

            if (VideoItems.Any(item => string.Equals(item.Name, lessonName, StringComparison.OrdinalIgnoreCase)))
            {
                duplicateCourses.Add(lessonName);
                AppendLog($"已跳过重复视频项: {lessonName}");
                return (true, true);
            }

            return (false, false);
        }

        private static string ExtractVideoUrlFromJson(System.Text.Json.JsonElement data)
        {
            if (!data.TryGetProperty("lessonData", out var lessonData)) return string.Empty;
            if (!lessonData.TryGetProperty("fileList", out var fileList) ||
                fileList.ValueKind != System.Text.Json.JsonValueKind.Array) return string.Empty;

            string lastValidUrl = string.Empty;
            foreach (var file in fileList.EnumerateArray())
            {
                if (file.TryGetProperty("Playset", out var playset) &&
                    playset.ValueKind == System.Text.Json.JsonValueKind.Array)
                {
                    foreach (var play in playset.EnumerateArray())
                    {
                        if (play.TryGetProperty("Url", out var urlElement))
                        {
                            var url = urlElement.GetString()?.Replace("\\", "") ?? "";
                            if (url.Contains(".mp4", StringComparison.OrdinalIgnoreCase))
                            {
                                lastValidUrl = url;
                            }
                        }
                    }
                }
            }
            return lastValidUrl;
        }

        private async Task<long> GetFileSizeIfNeededAsync(string url)
        {
            if (string.IsNullOrEmpty(url)) return 0;

            try
            {
                return await _downloadService.GetFileSizeAsync(url);
            }
            catch
            {
                return 0;
            }
        }

        private async Task<bool> CreateAndAddVideoItemAsync(
            string lessonName, string url, long fileSize)
        {
            var newItem = new VideoItem
            {
                Name = lessonName,
                Url = url,
                FileSize = fileSize,
                Status = "解析完成",
                StatusColorHex = "#FF008000"
            };

            _uiService.Invoke(() =>
            {
                VideoItems.Add(newItem);
                UpdateSingleVideoIndex(newItem);
            });

            _duplicateDetectionService.AddToCache(newItem);

            return !string.IsNullOrEmpty(url);
        }

        private async Task<bool> UseOriginalLineParsingAsync(
            string content, List<string> duplicateFiles, List<string> duplicateCourses)
        {
            var lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var currentLessonName = ExtractLessonNameFromLines(lines);

            if (string.IsNullOrEmpty(currentLessonName))
            {
                return false;
            }

            var (isDuplicate, _) = await CheckForDuplicatesAsync(currentLessonName, duplicateFiles, duplicateCourses);
            if (isDuplicate)
            {
                return false;
            }

            var finalUrl = ExtractUrlFromLines(lines);
            long fileSize = await GetFileSizeIfNeededAsync(finalUrl);

            return await CreateAndAddVideoItemAsync(currentLessonName, finalUrl, fileSize);
        }

        private static string ExtractLessonNameFromLines(string[] lines)
        {
            foreach (var line in lines)
            {
                if (line.Contains("lessonName", StringComparison.OrdinalIgnoreCase))
                {
                    return ExtractValue(line, "lessonName");
                }
            }
            return string.Empty;
        }

        private static string ExtractUrlFromLines(string[] lines)
        {
            string finalUrl = string.Empty;
            bool playsetEncountered = false;
            bool inFileItem = false;

            foreach (var line in lines)
            {
                if (line.Contains('{') && !inFileItem)
                {
                    inFileItem = true;
                    playsetEncountered = false;
                }
                else if (line.Contains('}') && inFileItem)
                {
                    inFileItem = false;
                }

                if (inFileItem && line.Contains("Playset", StringComparison.OrdinalIgnoreCase))
                {
                    playsetEncountered = true;
                }

                if (inFileItem &&
                    line.Contains("url", StringComparison.OrdinalIgnoreCase) &&
                    line.Contains(".mp4", StringComparison.OrdinalIgnoreCase))
                {
                    var videoUrl = ExtractValue(line, "url").Replace("\\", "");
                    if (playsetEncountered)
                    {
                        finalUrl = videoUrl;
                    }
                }
            }
            return finalUrl;
        }

        private static string ExtractValue(string jsonLine, string key)
        {
            try
            {
                var keyIndex = jsonLine.IndexOf(key, StringComparison.OrdinalIgnoreCase);
                if (keyIndex < 0) return string.Empty;

                var colonIndex = jsonLine.IndexOf(':', keyIndex + key.Length);
                if (colonIndex < 0) return string.Empty;

                var startIndex = colonIndex + 1;
                while (startIndex < jsonLine.Length && char.IsWhiteSpace(jsonLine[startIndex]))
                {
                    startIndex++;
                }

                var endIndex = startIndex;
                if (startIndex < jsonLine.Length)
                {
                    var startChar = jsonLine[startIndex];
                    var endChar = startChar == '"' ? '"' : ',';
                    endIndex = startChar == '"'
                        ? jsonLine.IndexOf(endChar, startIndex + 1)
                        : jsonLine.IndexOfAny(new[] { ',', '}', ']' }, startIndex);
                }

                if (endIndex < 0) endIndex = jsonLine.Length;

                return jsonLine.Substring(startIndex, endIndex - startIndex)
                    .Trim()
                    .Trim('"', '\'', ',', ' ');
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"提取值失败: {ex.Message}");
                return string.Empty;
            }
        }

        private void ShowDuplicateMessage(List<string> duplicateFiles, List<string> duplicateCourses)
        {
            if (duplicateFiles.Count == 0 && duplicateCourses.Count == 0) return;

            var message = new System.Text.StringBuilder();

            if (duplicateFiles.Count > 0)
            {
                message.AppendLine($"跳过 {duplicateFiles.Count} 个已存在的视频项：");
                foreach (var file in duplicateFiles.Take(5))
                {
                    message.AppendLine($"- {Truncate(file, 50)}");
                }
                if (duplicateFiles.Count > 5)
                {
                    message.AppendLine($"...及其他 {duplicateFiles.Count - 5} 个");
                }
                message.AppendLine();
            }

            if (duplicateCourses.Count > 0)
            {
                message.AppendLine($"跳过 {duplicateCourses.Count} 个重复的请求头：");
                foreach (var course in duplicateCourses.Take(5))
                {
                    message.AppendLine($"- {Truncate(course, 50)}");
                }
                if (duplicateCourses.Count > 5)
                {
                    message.AppendLine($"...及其他 {duplicateCourses.Count - 5} 个");
                }
            }

            if (message.Length > 0)
            {
                _uiService.Invoke(async () =>
                {
                    await _dialogService.ShowMessageBoxAsync(
                        message.ToString(),
                        "重复内容提示",
                        DialogButton.OK,
                        DialogIcon.Information);
                });
            }
        }

        private Task ExecuteDownloadAllCommandAsync()
        {
            var pendingItems = VideoItems
                .Where(item => item.StatusFlag != DownloadStatus.Completed &&
                              item.StatusFlag != DownloadStatus.Downloading)
                .ToList();

            if (pendingItems.Count == 0)
            {
                _loggingService.Info("用户尝试下载全部视频，但没有可下载的项目", "Download");
                AppendLog("没有可下载的视频");
                return Task.CompletedTask;
            }

            try
            {
                _loggingService.Info($"用户点击开始下载全部视频按钮，待下载项目数: {pendingItems.Count}", "UI.Command");
                AppendLog($"点击按钮：开始下载全部视频 (待下载 {pendingItems.Count} 个)");
                _loggingService.Info($"开始批量下载 {pendingItems.Count} 个视频项目", "Download");
                AppendLog($"开始下载视频...");
                
                foreach (var item in pendingItems)
                {
                    _loggingService.Debug($"添加下载任务: {Truncate(item.Name, 50)}", "Download.Queue");
                    _downloadManager.StartDownload(item);
                }
                
                _loggingService.Info("所有下载任务已添加到队列", "Download");
            }
            catch (Exception ex)
            {
                _loggingService.Error($"启动批量下载时发生异常: {ex.Message}", "Download");
                AppendLog($"启动下载失败: {ex.Message}");
            }

            return Task.CompletedTask;
        }

        private async void ExecuteOpenFolderCommand()
        {
            try
            {
                _loggingService.Info("用户点击打开下载目录按钮", "UI.Command");
                AppendLog("点击按钮：打开下载目录");
                
                if (Directory.Exists(DownloadPath))
                {
                    _loggingService.Debug($"尝试打开下载目录: {DownloadPath}", "FileSystem");
                    
                    // 在后台线程中启动进程，避免阻塞UI
                    await System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = "explorer.exe",
                                Arguments = DownloadPath,
                                UseShellExecute = true,
                                CreateNoWindow = true
                            });
                        }
                        catch (Exception ex)
                        {
                            _loggingService.Error($"启动explorer进程失败: {ex.Message}", "FileSystem");
                        }
                    });
                    
                    _loggingService.Info("下载目录已成功打开", "FileSystem");
                    AppendLog("正在打开下载目录...");
                    await System.Threading.Tasks.Task.Delay(1000);
                }
                else
                {
                    _loggingService.Warning($"下载目录不存在: {DownloadPath}", "FileSystem");
                    AppendLog("下载目录不存在");
                }
            }
            catch (Exception ex)
            {
                _loggingService.Error($"打开下载目录时发生异常: {ex.Message}", "FileSystem");
                AppendLog($"无法打开目录: {ex.Message}");
            }
        }

        private async Task ExecuteChangeDownloadPathCommandAsync()
        {
            try
            {
                _loggingService.Info("用户点击更改下载路径按钮", "UI.Command");
                var result = await _fileFolderService.SelectFolderAsync(
                    "选择下载目录",
                    Directory.Exists(DownloadPath) ? DownloadPath : "");

                if (!string.IsNullOrWhiteSpace(result))
                {
                    _loggingService.Info($"用户选择了新的下载路径: {result}", "FileSystem");
                    DownloadPath = result;
                    _loggingService.Info($"下载路径已成功更新为: {DownloadPath}", "Configuration");
                    AppendLog($"下载路径已更新为: {DownloadPath}");
                }
                else
                {
                    _loggingService.Debug("用户取消了目录选择操作", "FileSystem");
                }
            }
            catch (Exception ex)
            {
                _loggingService.Error($"更改下载路径时发生异常: {ex.Message}", "FileSystem");
                AppendLog($"更改目录失败: {ex.Message}");
            }
        }

        private Task ExecuteInitializeToolCommandAsync()
        {
            _loggingService.Info("用户点击初始化解析工具按钮", "UI.Command");
            AppendLog("点击按钮：初始化解析工具");
            
            var hasPendingItems = VideoItems.Any(item =>
                item.StatusFlag != DownloadStatus.Completed &&
                item.StatusFlag != DownloadStatus.Failed);

            if (hasPendingItems)
            {
                _loggingService.Warning("检测到未完成的下载任务，需要用户确认", "Initialization");
                var confirm = _dialogService.ShowConfirmDialog(
                    "当前有视频未完成下载或正在下载中，继续初始化将取消所有下载任务并清除视频列表。是否继续初始化？",
                    "警告");

                if (!confirm)
                {
                    _loggingService.Info("用户取消了初始化操作", "Initialization");
                    AppendLog("用户取消了初始化操作");
                    return Task.CompletedTask;
                }
                else
                {
                    _loggingService.Info("用户确认继续初始化，将取消所有下载任务", "Initialization");
                }
            }
            else
            {
                _loggingService.Debug("没有未完成的下载任务，直接进行初始化", "Initialization");
            }

            try
            {
                _loggingService.Info("开始执行工具初始化操作", "Initialization");
                _downloadManager.ClearAll();
                _logViewModel.ClearLog();
                
                DownloadPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, AppConstants.DefaultDownloadFolder);
                
                // 确保下载目录存在，添加错误处理
                try
                {
                    Directory.CreateDirectory(DownloadPath);
                    _loggingService.Info($"下载目录已创建: {DownloadPath}", "FileSystem");
                    AppendLog($"下载目录已重置: {DownloadPath}");
                }
                catch (Exception ex)
                {
                    _loggingService.Error($"创建下载目录失败: {ex.Message}", "FileSystem");
                    AppendLog($"创建下载目录失败: {ex.Message}");
                    // 即使目录创建失败，也继续初始化
                }
                
                _downloadManager.Initialize(DownloadPath);
                _healthMonitor.Reset();
                
                _loggingService.Info("解析工具初始化完成", "Initialization");
                AppendLog("解析工具初始化完成");
            }
            catch (Exception ex)
            {
                _loggingService.Fatal($"工具初始化过程中发生严重错误: {ex.Message}", "Initialization");
                AppendLog($"初始化失败: {ex.Message}");
            }

            return Task.CompletedTask;
        }

        private async void ExecuteCopyUrlCommand(VideoItem? item)
        {
            if (item == null) return;

            AppendLog($"点击按钮：复制视频URL ({Truncate(item.Name, 30)})");
            await _clipboardService.SetClipboardTextAsync(item.Url);
            AppendLog("视频URL已复制到剪贴板");
        }

        private void ExecuteDownloadSingleCommand(VideoItem? item)
        {
            if (item == null) return;

            _loggingService.Info($"用户点击下载单个视频按钮: {Truncate(item.Name, 50)}", "UI.Command");
            
            if (item.StatusFlag == DownloadStatus.Completed)
            {
                _loggingService.Info($"视频已下载完成，无需重复下载: {Truncate(item.Name, 50)}", "Download");
                AppendLog("该视频已下载完成");
                return;
            }

            if (item.StatusFlag == DownloadStatus.Downloading)
            {
                _loggingService.Debug($"视频正在下载中，忽略重复下载请求: {Truncate(item.Name, 50)}", "Download");
                return;
            }

            _loggingService.Debug($"添加单个视频下载任务: {Truncate(item.Name, 50)}", "Download.Queue");
            _downloadManager.StartDownload(item);
        }

        private void ExecuteDeleteItemCommand(VideoItem? item)
        {
            if (item == null) return;

            _loggingService.Info($"用户点击删除视频项按钮: {Truncate(item.Name, 50)}", "UI.Command");
            
            var logPrefix = $"删除视频项 ({Truncate(item.Name, 30)}) - ";

            if (item.StatusFlag == DownloadStatus.Downloading)
            {
                _loggingService.Warning($"尝试删除正在下载的视频项: {Truncate(item.Name, 50)}，需要用户确认", "Download");
                var confirm = _dialogService.ShowConfirmDialog(
                    $"视频 '{Truncate(item.Name, 50)}' 正在下载中，删除将取消下载。是否继续删除？",
                    "警告");

                if (!confirm)
                {
                    _loggingService.Info($"用户取消了删除操作: {Truncate(item.Name, 50)}", "UI");
                    AppendLog($"{logPrefix}用户取消了删除操作");
                    return;
                }

                item.DownloadTokenSource?.Cancel();
                _loggingService.Info($"已取消下载任务: {Truncate(item.Name, 50)}", "Download");
                AppendLog($"{logPrefix}下载任务已取消");
            }

            if (item.StatusFlag == DownloadStatus.Completed)
            {
                _loggingService.Debug($"从已完成计数中移除视频项: {Truncate(item.Name, 50)}", "Download");
                AppendLog($"{logPrefix}从已完成计数中移除");
            }

            VideoItems.Remove(item);
            UpdateVideoIndexes();

            _duplicateDetectionService.RemoveFromCache(item.Name);
            _loggingService.Info($"视频项删除成功: {Truncate(item.Name, 50)}", "VideoList");
            AppendLog($"{logPrefix} 成功");
        }

        private void UpdateVideoIndexes()
        {
            for (int i = 0; i < VideoItems.Count; i++)
            {
                VideoItems[i].DisplayIndex = i + 1;
            }
        }

        private void UpdateSingleVideoIndex(VideoItem item)
        {
            var index = VideoItems.IndexOf(item);
            if (index >= 0)
            {
                item.DisplayIndex = index + 1;
            }
        }

        private void ExecuteSliderValueChangedCommand(double value)
        {
            // 确保值在有效范围内
            int intValue = (int)Math.Round(value);
            if (intValue < 1) intValue = 1;
            if (intValue > 16) intValue = 16;
            
            // 只有当值真正改变时才记录日志
            if (intValue != _maxConcurrentDownloads)
            {
                _loggingService.Info($"用户调整最大并发下载数: {intValue}", "UI.Slider");
                AppendLogForSlider($"滑块值已更改: {intValue}");
                
                // 同步更新属性以触发UI更新
                MaxConcurrentDownloads = intValue;
            }
        }

        private void AppendLog(string message)
        {
            _loggingService.Info(message, "UI");
            _logViewModel.AppendLog(message);
            // 强制通知UI更新
            OnPropertyChanged(nameof(LogText));
        }

        public void AppendLogForSlider(string message)
        {
            _loggingService.Trace(message, "UI.Slider");
            _logViewModel.AppendLogForSlider(message);
        }

        private static string Truncate(string text, int maxLength)
        {
            return string.IsNullOrEmpty(text) || text.Length <= maxLength
                ? text
                : $"{text[..Math.Min(maxLength, text.Length)]}...";
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            _downloadManager.PropertyChanged -= OnDownloadManagerPropertyChanged;
            _downloadManager.Dispose();
            _logViewModel.Dispose();
        }
    }
}
