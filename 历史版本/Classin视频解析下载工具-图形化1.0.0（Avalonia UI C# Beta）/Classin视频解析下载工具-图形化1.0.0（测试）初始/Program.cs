using System;
using System.IO;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Microsoft.Extensions.DependencyInjection;
using VideoDownloader;
using VideoDownloader.Services;
using VideoDownloader.ViewModels;

namespace Classin视频解析下载工具
{
    internal sealed class Program
    {
        // Initialization code. Don't use any Avalonia, third-party APIs or any
        // SynchronizationContext-reliant code before AppMain is called: things aren't initialized
        // yet and stuff might break.
        [STAThread]
        public static void Main(string[] args)
        {
            // 创建日志文件
            var logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "startup.log");
            using (var writer = new StreamWriter(logFile))
            {
                writer.WriteLine($"应用程序开始启动: {DateTime.Now}");
                try
                {
                    writer.WriteLine("构建Avalonia应用...");
                    var appBuilder = BuildAvaloniaApp();
                    writer.WriteLine("启动应用程序...");
                    appBuilder.StartWithClassicDesktopLifetime(args);
                    writer.WriteLine($"应用程序正常退出: {DateTime.Now}");
                }
                catch (Exception ex)
                {
                    writer.WriteLine($"应用程序启动失败: {ex.Message}");
                    writer.WriteLine(ex.StackTrace);
                    Console.WriteLine($"应用程序启动失败: {ex.Message}");
                    Console.WriteLine(ex.StackTrace);
                    Environment.Exit(1);
                }
            }
        }

        // Avalonia configuration, don't remove; also used by visual designer.
        public static AppBuilder BuildAvaloniaApp()
            => AppBuilder.Configure<App>()
                .UsePlatformDetect()
                .WithInterFont()
                .LogToTrace();
    }
}
