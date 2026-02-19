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
        private bool _isDisposed;
        
        public override void Initialize()
        {
            Console.WriteLine("开始初始化应用程序...");
            try
            {
                Console.WriteLine("加载Avalonia XAML...");
                AvaloniaXamlLoader.Load(this);
                Console.WriteLine("配置服务...");
                ConfigureServices();
                Console.WriteLine("订阅全局异常...");
                SubscribeToGlobalExceptions();
                Console.WriteLine("应用程序初始化完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"初始化失败: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                ShowFatalErrorDialog($"初始化失败: {ex.Message}");
            }
        }

        private void SubscribeToGlobalExceptions()
        {
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        }

        private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;
            var message = exception?.Message ?? "未知错误";

            ShowFatalErrorDialog(message);

            CleanupResources();
        }

        private void ShowFatalErrorDialog(string message)
        {
            try
            {
                if (_serviceProvider != null)
                {
                    var dialogService = _serviceProvider.GetService(typeof(IDialogService)) as IDialogService;
                    dialogService?.ShowMessageBoxAsync(message, "严重错误", DialogButton.OK, DialogIcon.Error).Wait();
                }
            }
            catch
            {
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
            if (_isDisposed) return;
            _isDisposed = true;

            UnsubscribeFromGlobalExceptions();

            // 清理服务资源
            if (_serviceProvider != null)
                {
                    var networkMonitor = _serviceProvider.GetService(typeof(INetworkMonitor)) as INetworkMonitor;
                    if (networkMonitor is IDisposable disposable) {
                        disposable.Dispose();
                    }
                    
                    var loggingService = _serviceProvider.GetService(typeof(ILoggingService)) as ILoggingService;
                    if (loggingService is IDisposable disposable2) {
                        disposable2.Dispose();
                    }
                    
                    var downloadService = _serviceProvider.GetService(typeof(IDownloadService)) as IDownloadService;
                    if (downloadService is IDisposable disposable3) {
                        disposable3.Dispose();
                    }
                }
        }

        private void UnsubscribeFromGlobalExceptions()
        {
            AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
        }

        private void OnMainWindowClosed(object? sender, EventArgs e)
        {
            CleanupResources();
        }

        public override void OnFrameworkInitializationCompleted()
        {
            Console.WriteLine("开始初始化框架...");
            try
            {
                if (this.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
                {
                    Console.WriteLine("配置桌面应用程序生命周期...");
                    // Avoid duplicate validations from both Avalonia and the CommunityToolkit. 
                    // More info: https://docs.avaloniaui.net/docs/guides/development-guides/data-validation#manage-validationplugins
                    Console.WriteLine("禁用Avalonia数据注解验证...");
                    DisableAvaloniaDataAnnotationValidation();
                    
                    try
                    {
                        Console.WriteLine("获取MainViewModel...");
                        var mainViewModel = _serviceProvider!.GetService(typeof(MainViewModel)) as MainViewModel;
                        if (mainViewModel == null)
                        {
                            throw new InvalidOperationException("MainViewModel服务未注册");
                        }
                        var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                        Console.WriteLine($"初始化MainViewModel，基础目录: {baseDirectory}");
                        mainViewModel.Initialize(baseDirectory);
                        
                        Console.WriteLine("创建MainWindow...");
                        var mainWindow = new MainWindow
                        {
                            DataContext = mainViewModel,
                        };
                        
                        // 设置主窗口到服务中
                        Console.WriteLine("设置主窗口到服务中...");
                        var windowService = _serviceProvider!.GetService(typeof(IWindowService)) as IWindowService;
                        if (windowService == null)
                        {
                            throw new InvalidOperationException("IWindowService服务未注册");
                        }
                        windowService.SetMainWindow(mainWindow);
                        
                        Console.WriteLine("设置主窗口到DialogService中...");
                        var dialogService = _serviceProvider!.GetService(typeof(IDialogService)) as DialogService;
                        if (dialogService == null)
                        {
                            throw new InvalidOperationException("IDialogService服务未注册");
                        }
                        // 这里需要反射设置主窗口，因为构造函数参数问题
                        var mainWindowField = typeof(DialogService).GetField("_mainWindow", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        mainWindowField?.SetValue(dialogService, mainWindow);
                        
                        Console.WriteLine("设置主窗口到FileFolderService中...");
                        var fileFolderService = _serviceProvider!.GetService(typeof(IFileFolderService)) as FileFolderService;
                        if (fileFolderService == null)
                        {
                            throw new InvalidOperationException("IFileFolderService服务未注册");
                        }
                        {
                            var mainWindowField2 = typeof(FileFolderService).GetField("_mainWindow", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                            mainWindowField2?.SetValue(fileFolderService, mainWindow);
                        }
                        
                        Console.WriteLine("添加窗口关闭事件处理...");
                        mainWindow.Closing += (sender, e) => OnMainWindowClosed(sender, EventArgs.Empty);
                        Console.WriteLine("设置桌面主窗口...");
                        desktop.MainWindow = mainWindow;
                        Console.WriteLine("主窗口已设置完成...");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error creating main window: {ex.Message}");
                        Console.WriteLine(ex.StackTrace);
                        ShowFatalErrorDialog($"创建主窗口失败: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("ApplicationLifetime不是IClassicDesktopStyleApplicationLifetime");
                }

                base.OnFrameworkInitializationCompleted();
                Console.WriteLine("框架初始化完成...");
            }
            catch (Exception ex)
            {
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