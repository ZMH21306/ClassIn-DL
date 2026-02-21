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
        private readonly IMemoryMonitor _memoryMonitor;
        private readonly LogViewModel _logViewModel;
        private readonly DownloadManager _downloadManager;
        
        private string _downloadPath = string.Empty;
        private string _statusSummary = "当前没有视频正在下载，统计信息为空";
        private bool _isInitialized;
        private bool _disposed;

        public MainViewModel()
        {
            // 默认构造函数用于设计时支持
            _clipboardService = new ClipboardService();
            _validationService = new ValidationService();
            _configurationManager = new ConfigurationManager();
            _healthMonitor = new HealthMonitor();
            _downloadService = new DownloadService(_validationService);
            _uiService = new UIService();
            _fileFolderService = new FileFolderService();
            _dialogService = new DialogService();
            _loggingService = new LoggingService();
            _duplicateDetectionService = new DuplicateDetectionService(_loggingService);
            _parseService = new ParseService(_validationService, _duplicateDetectionService, _downloadService, _loggingService);
            _memoryMonitor = new MemoryMonitor();
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
            IMemoryMonitor memoryMonitor,
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
            _memoryMonitor = memoryMonitor;
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
        

        // 并行下载功能已暂时移除
        // public Classin视频解析下载工具.Commands.RelayCommand<double> SliderValueChangedCommand { get; private set; } = null!;

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
        }

        private async Task LoadConfigurationAsync()
        {
            try
            {
                // 使用异步方式加载配置，避免阻塞UI线程
                var config = await _configurationManager.GetConfigurationAsync();
                DownloadPath = config.DownloadPath;
            }
            catch (Exception ex)
            {
                _loggingService.Error($"加载配置失败: {ex.Message}", "Configuration");
                // 使用默认下载路径
                DownloadPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, AppConstants.DefaultDownloadFolder);
            }
        }
        
        private void LoadConfiguration()
        {
            // 同步版本，用于设计时支持
            try
            {
                var config = _configurationManager.GetConfiguration();
                DownloadPath = config.DownloadPath;
            }
            catch (Exception ex)
            {
                _loggingService?.Error($"加载配置失败: {ex.Message}", "Configuration");
                // 使用默认下载路径
                DownloadPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, AppConstants.DefaultDownloadFolder);
            }
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

        public async Task InitializeAsync(string baseDirectory)
        {
            _healthMonitor.RecordInitialization();
            
            // 启动内存监控
            _memoryMonitor.StartMonitoring();
            _memoryMonitor.TakeSnapshot("初始化开始");

            try
            {
                // 异步加载配置
                await LoadConfigurationAsync();
            }
            catch (Exception ex)
            {
                _loggingService.Error($"异步加载配置失败: {ex.Message}", "Configuration");
            }

            // 只有当下载路径为空时，才设置默认路径
            if (string.IsNullOrEmpty(DownloadPath))
            {
                var newDownloadPath = Path.Combine(baseDirectory, AppConstants.DefaultDownloadFolder);
                DownloadPath = newDownloadPath;
            }
            
            // 异步确保下载目录存在，添加错误处理
            try
            {
                await Task.Run(() => Directory.CreateDirectory(DownloadPath));
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

            // 异步初始化下载管理器
            await Task.Run(() => _downloadManager.Initialize(DownloadPath));
            
            // 拍摄初始化完成后的内存快照
            _memoryMonitor.TakeSnapshot("初始化完成");
            
            // 记录初始内存使用情况
            var initialMemory = _memoryMonitor.GetCurrentMemoryUsage();
            AppendLog($"内存使用情况: {initialMemory.PrivateMemorySizeMb:F2} MB");
                        
            IsInitialized = true;
            AppendLog("软件初始化完成");
        }
        
        public void Initialize(string baseDirectory)
        {
            // 同步版本，用于设计时支持
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
                
                // 拍摄解析前的内存快照
                _memoryMonitor.TakeSnapshot("解析开始");
                
                // 使用ParseService进行解析
                var duplicateFiles = new List<string>();
                var duplicateCourses = new List<string>();
                var (parseSuccess, duplicateFound, lessonName, videoUrl) = await _parseService.ParseContentAsync(clipboardText, duplicateFiles, duplicateCourses);

                // 拍摄解析后的内存快照
                _memoryMonitor.TakeSnapshot("解析完成");
                var memoryUsage = _memoryMonitor.GetCurrentMemoryUsage();
                AppendLog($"解析完成，内存使用: {memoryUsage.PrivateMemorySizeMb:F2} MB");

                if (parseSuccess && !string.IsNullOrEmpty(lessonName) && !string.IsNullOrEmpty(videoUrl))
                {
                    // 获取文件大小
                    long fileSize = await GetFileSizeIfNeededAsync(videoUrl);
                    
                    // 创建并添加视频项
                    await CreateAndAddVideoItemAsync(lessonName, videoUrl, fileSize);
                    
                    _loggingService.Info("请求头解析成功完成", "Parse");
                    AppendLog("请求头解析成功");
                }
                else
                {
                    _loggingService.Warning("请求头解析失败", "Parse");
                    AppendLog("请求头解析失败");
                }

                // 显示重复消息
                ShowDuplicateMessage(duplicateFiles, duplicateCourses);
            }
            catch (Exception ex)
            {
                _loggingService.Error($"解析请求头时发生异常: {ex.Message}", "Parse");
                AppendLog($"解析失败: {ex.Message}");
            }
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
            await Task.CompletedTask;
            
            var newItem = new VideoItem
            {
                Name = lessonName,
                Url = url,
                FileSize = fileSize,
                Status = "解析完成",
                StatusColorHex = "#FF0000FF"
            };

            _uiService.Invoke(() =>
            {
                _downloadManager.AddVideoItem(newItem);
            });

            _duplicateDetectionService.AddToCache(newItem);

            return !string.IsNullOrEmpty(url);
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

        private async Task ExecuteDownloadAllCommandAsync()
        {
            var pendingItems = VideoItems
                .Where(item => item.StatusFlag != DownloadStatus.Completed &&
                              item.StatusFlag != DownloadStatus.Downloading)
                .ToList();

            if (pendingItems.Count == 0)
            {
                _loggingService.Info("用户尝试下载全部视频，但没有可下载的项目", "Download");
                AppendLog("没有可下载的视频");
                return;
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

            return;
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

        private async Task ExecuteInitializeToolCommandAsync()
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
                    return;
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
                    await Task.Run(() => Directory.CreateDirectory(DownloadPath));
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

            return;
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

        private void AppendLog(string message)
        {
            _loggingService.Info(message, "UI");
            _logViewModel.AppendLog(message);
            // 强制通知UI更新
            OnPropertyChanged(nameof(LogText));
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

            // 取消订阅事件
            _downloadManager.PropertyChanged -= OnDownloadManagerPropertyChanged;

            // 拍摄最终内存快照
            _memoryMonitor.TakeSnapshot("程序退出");
            
            // 检查是否有内存泄漏
            if (_memoryMonitor.IsMemoryLeakDetected())
            {
                AppendLog("警告: 检测到可能的内存泄漏");
                double memoryGrowth = _memoryMonitor.CalculateMemoryGrowth();
                AppendLog($"内存增长: {memoryGrowth:F2} MB");
            }
            else
            {
                var finalMemory = _memoryMonitor.GetCurrentMemoryUsage();
                AppendLog($"程序退出，最终内存使用: {finalMemory.PrivateMemorySizeMb:F2} MB");
            }

            // 停止内存监控
            _memoryMonitor.StopMonitoring();

            // 释放实现IDisposable的服务
            _downloadManager.Dispose();
            _logViewModel.Dispose();
            
            // 释放其他可能的可释放资源
            if (_downloadService is IDisposable downloadServiceDisposable)
            {
                downloadServiceDisposable.Dispose();
            }
            
            if (_loggingService is IDisposable loggingServiceDisposable)
            {
                loggingServiceDisposable.Dispose();
            }
            
            if (_memoryMonitor is IDisposable memoryMonitorDisposable)
            {
                memoryMonitorDisposable.Dispose();
            }
        }
    }
}
