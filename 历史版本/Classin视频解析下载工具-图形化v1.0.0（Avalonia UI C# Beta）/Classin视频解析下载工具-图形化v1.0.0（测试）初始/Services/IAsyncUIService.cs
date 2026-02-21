using System;
using System.Threading.Tasks;

namespace VideoDownloader.Services
{
    public interface IAsyncUIService
    {
        Task InvokeAsync(Action action);
        Task<T> InvokeAsync<T>(Func<T> func);
        Task<T> InvokeAsync<T>(Func<Task<T>> func);
        Task InvokeAsync(Func<Task> func);
    }
}
