using System;
using System.Threading.Tasks;
using Avalonia.Threading;

namespace Classin视频解析下载工具.Services
{
    public class UIService : IUIService, IAsyncUIService
    {
        public void Invoke(Action action)
        {
            if (Dispatcher.UIThread.CheckAccess())
            {
                action?.Invoke();
            }
            else
            {
                Dispatcher.UIThread.Post(action);
            }
        }

        public async Task InvokeAsync(Action action)
        {
            if (Dispatcher.UIThread.CheckAccess())
            {
                action?.Invoke();
                await Task.CompletedTask;
            }
            else
            {
                var tcs = new TaskCompletionSource<object?>();
                Dispatcher.UIThread.Post(() =>
                {
                    try
                    {
                        action?.Invoke();
                        tcs.SetResult(null);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                });
                await tcs.Task;
            }
        }

        public async Task<T> InvokeAsync<T>(Func<T> func)
        {
            if (Dispatcher.UIThread.CheckAccess())
            {
                return func();
            }
            else
            {
                var tcs = new TaskCompletionSource<T>();
                Dispatcher.UIThread.Post(() =>
                {
                    try
                    {
                        tcs.SetResult(func());
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                });
                return await tcs.Task;
            }
        }

        public async Task<T> InvokeAsync<T>(Func<Task<T>> func)
        {
            if (Dispatcher.UIThread.CheckAccess())
            {
                return await func();
            }
            else
            {
                var tcs = new TaskCompletionSource<T>();
                Dispatcher.UIThread.Post(async () =>
                {
                    try
                    {
                        var result = await func();
                        tcs.SetResult(result);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                });
                return await tcs.Task;
            }
        }

        public async Task InvokeAsync(Func<Task> func)
        {
            if (Dispatcher.UIThread.CheckAccess())
            {
                await func();
            }
            else
            {
                var tcs = new TaskCompletionSource<object?>();
                Dispatcher.UIThread.Post(async () =>
                {
                    try
                    {
                        await func();
                        tcs.SetResult(null);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                });
                await tcs.Task;
            }
        }

        public void ShowMessage(string message, string title = "信息")
        {
            Invoke(() =>
            {
                Console.WriteLine($"{title}: {message}");
            });
        }

        public void ShowError(string message, string title = "错误")
        {
            Invoke(() =>
            {
                Console.WriteLine($"{title}: {message}");
            });
        }

        public bool ShowConfirm(string message, string title = "确认")
        {
            var result = false;
            Invoke(() =>
            {
                Console.WriteLine($"{title}: {message}");
                result = true;
            });
            return result;
        }
    }
}