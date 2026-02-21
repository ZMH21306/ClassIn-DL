using System;
using System.Linq;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Data.Core;
using Avalonia.Data.Core.Plugins;
using Avalonia.Markup.Xaml;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Classin视频解析下载工具.Views;
using Classin视频解析下载工具.Services;
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
                _loggingService?.Info("开始初始化应用程序...", "Application");
                
                _loggingService?.Debug("加载Avalonia XAML...", "Application.Startup");
                AvaloniaXamlLoader.Load(this);
                
                _loggingService?.Debug("配置服务...", "Application.Startup");
                ConfigureServices();
                
                _loggingService?.Debug("订阅全局异常...", "Application.Startup");
                SubscribeToGlobalExceptions();
                
                _loggingService?.Info("应用程序初始化完成", "Application");
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
                _loggingService?.Debug("清理NetworkMonitor资源...", "Application.Shutdown");
                var networkMonitor = _serviceProvider.GetService(typeof(INetworkMonitor)) as INetworkMonitor;
                if (networkMonitor is IDisposable disposable) 
                {
                    disposable.Dispose();
                }
                
                _loggingService?.Debug("清理DownloadService资源...", "Application.Shutdown");
                var downloadService = _serviceProvider.GetService(typeof(IDownloadService)) as IDownloadService;
                if (downloadService is IDisposable disposable3) 
                {
                    disposable3.Dispose();
                }
                
                // 注意：最后才清理LoggingService，因为它还在被使用
                _loggingService?.Debug("清理LoggingService资源...", "Application.Shutdown");
                var loggingService = _serviceProvider.GetService(typeof(ILoggingService)) as ILoggingService;
                if (loggingService is IDisposable disposable2) 
                {
                    disposable2.Dispose();
                }
            }
            
            _loggingService?.Info("应用程序资源清理完成", "Application.Shutdown");
        }

        private void UnsubscribeFromGlobalExceptions()
        {
            _loggingService?.Debug("取消订阅全局异常处理", "Application.Exception");
            AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
        }

        private void OnMainWindowClosed(object? sender, EventArgs e)
        {
            _loggingService?.Info("主窗口已关闭，开始清理资源...", "Application.Shutdown");
            CleanupResources();
        }

        public override void OnFrameworkInitializationCompleted()
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
                        mainViewModel.Initialize(baseDirectory);
                        
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
                        mainWindow.Closing += (sender, e) => OnMainWindowClosed(sender, EventArgs.Empty);
                        _loggingService?.Debug("设置桌面主窗口...", "Application.Framework");
                        desktop.MainWindow = mainWindow;
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

            // Services
            services.AddLogging();

            services.AddSingleton<ILoggingService, LoggingService>();
            services.AddSingleton<IValidationService, ValidationService>();
            services.AddSingleton<IDuplicateDetectionService, DuplicateDetectionService>();
            services.AddSingleton<IConfigurationManager, ConfigurationManager>();
            services.AddSingleton<IProgressTrackingService, ProgressTrackingService>();
            services.AddSingleton<IHealthMonitor, HealthMonitor>();
            services.AddSingleton<INetworkMonitor, NetworkMonitor>();
            services.AddSingleton<IUIService, UIService>();
            services.AddSingleton<IAsyncUIService, UIService>();
            services.AddSingleton<IClipboardService, ClipboardService>();
            services.AddSingleton<IFileFolderService, FileFolderService>();
            services.AddSingleton<IDialogService, DialogService>();
            services.AddSingleton<IWindowService, WindowService>();
            
            services.AddSingleton<IDownloadService>(sp =>
            {
                var validationService = sp.GetRequiredService<IValidationService>();
                return new DownloadService(validationService);
            });
            
            services.AddSingleton<IParseService>(sp =>
            {
                var validationService = sp.GetRequiredService<IValidationService>();
                var duplicateDetectionService = sp.GetRequiredService<IDuplicateDetectionService>();
                var downloadService = sp.GetRequiredService<IDownloadService>();
                return new ParseService(validationService, duplicateDetectionService, downloadService);
            });

            // ViewModels
            services.AddSingleton<ViewModels.LogViewModel>();
            services.AddSingleton<ViewModels.DownloadManager>();
            services.AddSingleton<ViewModels.MainViewModel>();
            services.AddSingleton<ViewModels.SettingsViewModel>();
            services.AddSingleton<ViewModels.VideoListViewModel>();

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