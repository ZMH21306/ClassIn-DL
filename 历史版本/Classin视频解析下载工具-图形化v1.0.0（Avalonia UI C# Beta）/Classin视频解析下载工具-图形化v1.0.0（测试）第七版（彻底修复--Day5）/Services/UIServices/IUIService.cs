using System;
using System.Threading.Tasks;

namespace Classin视频解析下载工具.Services.UIServices
{
    public interface IUIService
    {
        void Invoke(Action action);
        Task InvokeAsync(Action action);
        Task<T> InvokeAsync<T>(Func<T> func);
        void ShowMessage(string message, string title = "信息");
        void ShowError(string message, string title = "错误");
        bool ShowConfirm(string message, string title = "确认");
    }
}
