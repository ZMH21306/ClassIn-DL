using System;
using System.Threading.Tasks;

namespace VideoDownloader.Services
{
    public class UIService : IUIService, IAsyncUIService
    {
        public void Invoke(Action action)
        {
            // 在Avalonia中，直接执行即可
            action?.Invoke();
        }

        public async Task InvokeAsync(Action action)
        {
            await Task.Run(action);
        }

        public async Task<T> InvokeAsync<T>(Func<T> func)
        {
            return await Task.Run(func);
        }

        public async Task<T> InvokeAsync<T>(Func<Task<T>> func)
        {
            return await func();
        }

        public async Task InvokeAsync(Func<Task> func)
        {
            await func();
        }

        public void ShowMessage(string message, string title = "信息")
        {
            // 简单实现，实际应该使用DialogService
            Console.WriteLine($"{title}: {message}");
        }

        public void ShowError(string message, string title = "错误")
        {
            // 简单实现，实际应该使用DialogService
            Console.WriteLine($"{title}: {message}");
        }

        public bool ShowConfirm(string message, string title = "确认")
        {
            // 简单实现，实际应该使用DialogService
            Console.WriteLine($"{title}: {message}");
            return true;
        }
    }
}