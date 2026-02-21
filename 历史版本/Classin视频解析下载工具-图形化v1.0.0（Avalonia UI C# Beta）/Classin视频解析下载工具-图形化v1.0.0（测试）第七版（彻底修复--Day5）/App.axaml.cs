using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Data.Core;
using Avalonia.Data.Core.Plugins;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Classin视频解析下载工具.Views;
using Classin视频解析下载工具.Services.CoreServices;
using Classin视频解析下载工具.Services.DownloadServices;
using Classin视频解析下载工具.Services.ParseServices;
using Classin视频解析下载工具.Services.SystemServices;
using Classin视频解析下载工具.Services.UIServices;
using Classin视频解析下载工具.ViewModels;

namespace Classin视频解析下载工具
{
    public partial class App : Application
    {
        private IServiceProvider? _serviceProvider;
        private ILoggingService? _loggingService;
        private bool _isDisposed;

        public override void Initialize()
        {
            try
            {
                Console.WriteLine("开始初始化应用程序...");

                Console.WriteLine("加载Avalonia XAML...");
                AvaloniaXamlLoader.Load(this);
                Console.WriteLine("加载Avalonia XAML完成");

                Console.WriteLine("配置服务...");
                ConfigureServices();
                Console.WriteLine("配置服务完成");

                _loggingService?.Info("开始初始化应用程序...", "Application");
                _loggingService?.Debug("加载Avalonia XAML...", "Application.Startup");
                _loggingService?.Debug("配置服务...", "Application.Startup");

                _loggingService?.Debug("订阅全局异常...", "Application.Startup");
                SubscribeToGlobalExceptions();

                _loggingService?.Info("应用程序初始化完成", "Application");
                Console.WriteLine("应用程序初始化完成");
            }
            catch (Exception ex)
            {
                _loggingService?.Fatal($"应用程序初始化失败: {ex.Message}", "Application");
                Console.WriteLine($"初始化失败: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                ShowFatalErrorDialog($"初始化失败: {ex.Message}");
            }
        }

        private void SubscribeToGlobalExceptions()
        {
            _loggingService?.Debug("订阅应用程序域未处理异常", "Application.Exception");
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        }

        private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;
            var message = exception?.Message ?? "未知错误";

            _loggingService?.Fatal($"捕获到未处理的应用程序域异常: {message}", "Application.Exception");
            if (exception != null)
            {
                _loggingService?.Fatal($"异常详情: {exception}", "Application.Exception");
            }

            ShowFatalErrorDialog(message);

            CleanupResources();
        }

        private void ShowFatalErrorDialog(string message)
        {
            _loggingService?.Fatal($"显示致命错误对话框: {message}", "Application.UI");

            try
            {
                if (_serviceProvider != null)
                {
                    var dialogService = _serviceProvider.GetService(typeof(IDialogService)) as IDialogService;
                    dialogService?.ShowMessageBoxAsync(message, "严重错误", DialogButton.OK, DialogIcon.Error).Wait();
                }
            }
            catch (Exception ex)
            {
                _loggingService?.Error($"显示错误对话框时发生异常: {ex.Message}", "Application.UI");
                // 降级到控制台输出
                Console.WriteLine($"严重错误: {message}\n\n应用程序即将退出。");
            }
        }

        private void ShowErrorDialog(string title, string message)
        {
            try
            {
                if (_serviceProvider != null)
                {
                    var dialogService = _serviceProvider.GetService(typeof(IDialogService)) as IDialogService;
                    dialogService?.ShowMessageBoxAsync(message, title, DialogButton.OK, DialogIcon.Error).Wait();
                }
            }
            catch
            {
                // 降级到控制台输出
                Console.WriteLine($"{title}: {message}");
            }
        }

        private void CleanupResources()
        {
            if (_isDisposed)
            {
                _loggingService?.Debug("资源已清理，跳过重复清理", "Application.Shutdown");
                return;
            }

            _loggingService?.Info("开始清理应用程序资源...", "Application.Shutdown");
            _isDisposed = true;

            UnsubscribeFromGlobalExceptions();

            // 清理服务资源
            if (_serviceProvider != null)
            {
                try
                {
                    _loggingService?.Debug("清理NetworkMonitor资源...", "Application.Shutdown");
                    var networkMonitor = _serviceProvider.GetService(typeof(INetworkMonitor)) as INetworkMonitor;
                    if (networkMonitor is IDisposable disposable)
                    {
                        disposable.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    _loggingService?.Error($"清理NetworkMonitor失败: {ex.Message}", "Application.Shutdown");
                }

                try
                {
                    _loggingService?.Debug("清理DownloadService资源...", "Application.Shutdown");
                    var downloadService = _serviceProvider.GetService(typeof(IDownloadService)) as IDownloadService;
                    if (downloadService is IDisposable disposable3)
                    {
                        disposable3.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    _loggingService?.Error($"清理DownloadService失败: {ex.Message}", "Application.Shutdown");
                }

                try
                {
                    _loggingService?.Debug("清理ParseService资源...", "Application.Shutdown");
                    var parseService = _serviceProvider.GetService(typeof(IParseService)) as IParseService;
                    if (parseService is IDisposable disposable4)
                    {
                        disposable4.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    _loggingService?.Error($"清理ParseService失败: {ex.Message}", "Application.Shutdown");
                }

                try
                {
                    _loggingService?.Debug("清理其他服务资源...", "Application.Shutdown");
                    var dialogService = _serviceProvider.GetService(typeof(IDialogService)) as IDialogService;
                    if (dialogService is IDisposable disposable5)
                    {
                        disposable5.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    _loggingService?.Error($"清理DialogService失败: {ex.Message}", "Application.Shutdown");
                }

                try
                {
                    var windowService = _serviceProvider.GetService(typeof(IWindowService)) as IWindowService;
                    if (windowService is IDisposable disposable6)
                    {
                        disposable6.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    _loggingService?.Error($"清理WindowService失败: {ex.Message}", "Application.Shutdown");
                }

                // 注意：最后才清理LoggingService，因为它还在被使用
                try
                {
                    _loggingService?.Debug("清理LoggingService资源...", "Application.Shutdown");
                    var loggingService = _serviceProvider.GetService(typeof(ILoggingService)) as ILoggingService;
                    if (loggingService is IDisposable disposable2)
                    {
                        disposable2.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    // 这里无法使用loggingService，因为可能已经被清理
                    Console.WriteLine($"清理LoggingService失败: {ex.Message}");
                }
            }

            Console.WriteLine("应用程序资源清理完成");
        }

        private void UnsubscribeFromGlobalExceptions()
        {
            _loggingService?.Debug("取消订阅全局异常处理", "Application.Exception");
            AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
        }

        private void OnMainWindowClosing(object? sender, Avalonia.Controls.WindowClosingEventArgs e)
        {
            _loggingService?.Info("主窗口开始关闭，检查活跃下载任务...", "Application.Shutdown");

            try
            {
                if (_serviceProvider != null)
                {
                    var downloadService = _serviceProvider.GetService(typeof(IDownloadService)) as IDownloadService;
                    if (downloadService != null && downloadService.HasActiveDownloads())
                    {
                        _loggingService?.Info($"检测到 {downloadService.GetActiveDownloadCount()} 个活跃下载任务", "Application.Shutdown");

                        // 显示退出确认对话框
                        var dialogService = _serviceProvider.GetService(typeof(IDialogService)) as IDialogService;
                        if (dialogService != null)
                        {
                            var result = dialogService.ShowConfirmDialog(
                                $"当前有 {downloadService.GetActiveDownloadCount()} 个下载任务正在进行中，确定要退出吗？\n退出将取消所有正在进行的下载。",
                                "确认退出");

                            if (!result)
                            {
                                _loggingService?.Info("用户取消退出", "Application.Shutdown");
                                e.Cancel = true;
                                return;
                            }
                            else
                            {
                                _loggingService?.Info("用户确认退出，开始取消下载任务", "Application.Shutdown");
                                // 取消所有下载任务
                                downloadService.Stop();
                            }
                        }
                    }
                }

                _loggingService?.Info("开始清理资源...", "Application.Shutdown");
                CleanupResources();
            }
            catch (Exception ex)
            {
                _loggingService?.Error($"退出确认过程发生异常: {ex.Message}", "Application.Shutdown");
                Console.WriteLine($"退出确认异常: {ex.Message}");
                // 即使异常也继续关闭
                CleanupResources();
            }
        }

        private void OnMainWindowClosed(object? sender, EventArgs e)
        {
            _loggingService?.Info("主窗口已关闭", "Application.Shutdown");
        }

        public override async void OnFrameworkInitializationCompleted()
        {
            _loggingService?.Info("开始初始化框架...", "Application.Framework");

            try
            {
                if (this.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
                {
                    _loggingService?.Debug("配置桌面应用程序生命周期...", "Application.Framework");
                    // Avoid duplicate validations from both Avalonia and the CommunityToolkit. 
                    // More info: https://docs.avaloniaui.net/docs/guides/development-guides/data-validation#manage-validationplugins
                    _loggingService?.Debug("禁用Avalonia数据注解验证...", "Application.Framework");
                    DisableAvaloniaDataAnnotationValidation();

                    try
                    {
                        _loggingService?.Debug("获取MainViewModel...", "Application.Framework");
                        var mainViewModel = _serviceProvider!.GetService(typeof(MainViewModel)) as MainViewModel;
                        if (mainViewModel == null)
                        {
                            throw new InvalidOperationException("MainViewModel服务未注册");
                        }
                        var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                        _loggingService?.Debug($"初始化MainViewModel，基础目录: {baseDirectory}", "Application.Framework");

                        // 创建主窗口，先不设置为桌面主窗口
                        _loggingService?.Debug("创建MainWindow...", "Application.Framework");
                        var mainWindow = new MainWindow
                        {
                            DataContext = mainViewModel,
                        };

                        // 设置主窗口到服务中
                        _loggingService?.Debug("设置主窗口到服务中...", "Application.Framework");
                        var windowService = _serviceProvider!.GetService(typeof(IWindowService)) as IWindowService;
                        if (windowService == null)
                        {
                            throw new InvalidOperationException("IWindowService服务未注册");
                        }
                        windowService.SetMainWindow(mainWindow);

                        _loggingService?.Debug("设置主窗口到DialogService中...", "Application.Framework");
                        var dialogService = _serviceProvider!.GetService(typeof(IDialogService)) as IDialogService;
                        if (dialogService == null)
                        {
                            throw new InvalidOperationException("IDialogService服务未注册");
                        }
                        dialogService.SetMainWindow(mainWindow);

                        _loggingService?.Debug("设置主窗口到FileFolderService中...", "Application.Framework");
                        var fileFolderService = _serviceProvider!.GetService(typeof(IFileFolderService)) as IFileFolderService;
                        if (fileFolderService == null)
                        {
                            throw new InvalidOperationException("IFileFolderService服务未注册");
                        }
                        fileFolderService.SetMainWindow(mainWindow);

                        _loggingService?.Debug("添加窗口关闭事件处理...", "Application.Framework");
                        mainWindow.Closing += OnMainWindowClosing;

                        // 先设置桌面主窗口，让窗口显示出来
                        _loggingService?.Debug("设置桌面主窗口...", "Application.Framework");
                        desktop.MainWindow = mainWindow;

                        // 异步初始化MainViewModel，不阻塞UI线程
                        _loggingService?.Debug("异步初始化MainViewModel...", "Application.Framework");
                        await Task.Run(async () => await mainViewModel.InitializeAsync(baseDirectory));

                        _loggingService?.Info("主窗口已设置完成...", "Application.Framework");
                    }
                    catch (Exception ex)
                    {
                        _loggingService?.Fatal($"创建主窗口失败: {ex.Message}", "Application.Framework");
                        Console.WriteLine($"Error creating main window: {ex.Message}");
                        Console.WriteLine(ex.StackTrace);
                        ShowFatalErrorDialog($"创建主窗口失败: {ex.Message}");
                    }
                }
                else
                {
                    _loggingService?.Warning("ApplicationLifetime不是IClassicDesktopStyleApplicationLifetime", "Application.Framework");
                    Console.WriteLine("ApplicationLifetime不是IClassicDesktopStyleApplicationLifetime");
                }

                base.OnFrameworkInitializationCompleted();
                _loggingService?.Info("框架初始化完成...", "Application.Framework");
            }
            catch (Exception ex)
            {
                _loggingService?.Fatal($"框架初始化失败: {ex.Message}", "Application.Framework");
                Console.WriteLine($"框架初始化失败: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                ShowFatalErrorDialog($"框架初始化失败: {ex.Message}");
            }
        }

        private void ConfigureServices()
        {
            var services = new ServiceCollection();

            // 核心服务 - 立即初始化
            services.AddSingleton<ILoggingService, LoggingService>();
            services.AddSingleton<IValidationService, ValidationService>();
            services.AddSingleton<IUIService, UIService>();
            services.AddSingleton<IAsyncUIService, UIService>();

            // 基础服务 - 延迟加载
            services.AddSingleton<IDuplicateDetectionService, DuplicateDetectionService>();
            services.AddSingleton<IConfigurationManager, ConfigurationManager>();
            services.AddSingleton<IProgressTrackingService, ProgressTrackingService>();
            services.AddSingleton<IHealthMonitor, HealthMonitor>();
            services.AddSingleton<IMemoryMonitor, MemoryMonitor>();

            // 网络相关服务 - 延迟加载
            services.AddSingleton<INetworkMonitor, NetworkMonitor>();

            // 剪贴板服务 - 延迟加载
            services.AddSingleton<IClipboardService, ClipboardService>();

            // 文件和对话框服务 - 延迟加载
            services.AddSingleton<IFileFolderService, FileFolderService>();
            services.AddSingleton<IDialogService, DialogService>();
            services.AddSingleton<IWindowService, WindowService>();

            // 下载服务 - 依赖ValidationService
            services.AddSingleton<IDownloadService>(sp =>
            {
                var validationService = sp.GetRequiredService<IValidationService>();
                return new DownloadService(validationService);
            });

            // 解析服务 - 依赖ValidationService, DuplicateDetectionService, DownloadService
            services.AddSingleton<IParseService>(sp =>
            {
                var validationService = sp.GetRequiredService<IValidationService>();
                var duplicateDetectionService = sp.GetRequiredService<IDuplicateDetectionService>();
                var downloadService = sp.GetRequiredService<IDownloadService>();
                var loggingService = sp.GetRequiredService<ILoggingService>();
                return new ParseService(validationService, duplicateDetectionService, downloadService, loggingService);
            });

            // ViewModels - 延迟加载
            services.AddSingleton<ViewModels.LogViewModel>(sp =>
            {
                var loggingService = sp.GetRequiredService<ILoggingService>();
                return new ViewModels.LogViewModel(loggingService);
            });

            services.AddSingleton<Classin视频解析下载工具.Services.DownloadServices.DownloadManager>(sp =>
            {
                var downloadService = sp.GetRequiredService<IDownloadService>();
                var duplicateDetectionService = sp.GetRequiredService<IDuplicateDetectionService>();
                var progressTrackingService = sp.GetRequiredService<IProgressTrackingService>();
                var loggingService = sp.GetRequiredService<ILoggingService>();
                return new Classin视频解析下载工具.Services.DownloadServices.DownloadManager(downloadService, duplicateDetectionService, progressTrackingService, loggingService);
            });

            services.AddSingleton<ViewModels.MainViewModel>(sp =>
            {
                var clipboardService = sp.GetRequiredService<IClipboardService>();
                var validationService = sp.GetRequiredService<IValidationService>();
                var duplicateDetectionService = sp.GetRequiredService<IDuplicateDetectionService>();
                var configurationManager = sp.GetRequiredService<IConfigurationManager>();
                var healthMonitor = sp.GetRequiredService<IHealthMonitor>();
                var downloadService = sp.GetRequiredService<IDownloadService>();
                var uiService = sp.GetRequiredService<IUIService>();
                var fileFolderService = sp.GetRequiredService<IFileFolderService>();
                var dialogService = sp.GetRequiredService<IDialogService>();
                var parseService = sp.GetRequiredService<IParseService>();
                var loggingService = sp.GetRequiredService<ILoggingService>();
                var memoryMonitor = sp.GetRequiredService<IMemoryMonitor>();
                var logViewModel = sp.GetRequiredService<ViewModels.LogViewModel>();
                var downloadManager = sp.GetRequiredService<Classin视频解析下载工具.Services.DownloadServices.DownloadManager>();
                return new ViewModels.MainViewModel(clipboardService, validationService, duplicateDetectionService, configurationManager, healthMonitor, downloadService, uiService, fileFolderService, dialogService, parseService, loggingService, memoryMonitor, logViewModel, downloadManager);
            });

            services.AddSingleton<ViewModels.SettingsViewModel>();
            services.AddSingleton<ViewModels.VideoListViewModel>();

            // 构建服务提供者
            _serviceProvider = services.BuildServiceProvider();

            // 获取LoggingService实例
            _loggingService = _serviceProvider.GetService<ILoggingService>();
        }

        private void DisableAvaloniaDataAnnotationValidation()
        {
            // 在Avalonia 11+中，我们需要使用不同的方式来管理验证插件
            try
            {
                // 直接使用Avalonia.Data.Core.Plugins命名空间中的类型
                var dataValidators = Avalonia.Data.Core.Plugins.BindingPlugins.DataValidators;
                if (dataValidators != null)
                {
                    // 移除所有数据注解验证插件
                    var pluginsToRemove = dataValidators.Where(p => p.GetType().Name.Contains("DataAnnotationsValidation")).ToArray();
                    foreach (var plugin in pluginsToRemove)
                    {
                        dataValidators.Remove(plugin);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"禁用数据注解验证时出错: {ex.Message}");
            }
        }
    }
}