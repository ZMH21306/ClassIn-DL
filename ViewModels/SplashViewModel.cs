using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Classin视频解析下载工具.Services.CoreServices;

namespace Classin视频解析下载工具.ViewModels
{
    public class SplashViewModel : ViewModelBase, IDisposable
    {
        private readonly ILoggingService _loggingService;
        private int _progressValue;
        private string _statusText = "正在初始化...";
        private bool _disposed;

        public SplashViewModel(ILoggingService loggingService)
        {
            _loggingService = loggingService;
            _loggingService.Info("启动画面初始化", "Splash");
        }

        public int ProgressValue
        {
            get => _progressValue;
            set => SetProperty(ref _progressValue, value);
        }

        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }

        public async Task InitializeAsync()
        {
            var steps = new[]
            {
                ("加载配置", 20),
                ("初始化下载引擎", 40),
                ("准备就绪", 70),
                ("完成", 100)
            };

            foreach (var (step, progress) in steps)
            {
                if (_disposed) break;
                StatusText = $"正在{step}...";
                ProgressValue = progress;
                _loggingService.Info($"启动中: {step}", "Splash");
                await Task.Delay(400);
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
        }
    }
}
