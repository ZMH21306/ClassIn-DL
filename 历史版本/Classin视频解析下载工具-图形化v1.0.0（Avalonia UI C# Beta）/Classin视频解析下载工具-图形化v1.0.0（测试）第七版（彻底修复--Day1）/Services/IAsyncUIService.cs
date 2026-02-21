using System;
using System.Threading.Tasks;

namespace Classin视频解析下载工具.Services
{
    public interface IAsyncUIService
    {
        Task InvokeAsync(Action action);
        Task<T> InvokeAsync<T>(Func<T> func);
        Task<T> InvokeAsync<T>(Func<Task<T>> func);
        Task InvokeAsync(Func<Task> func);
    }
}
