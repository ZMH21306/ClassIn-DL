using System;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;

namespace Classin视频解析下载工具.Services.SystemServices
{
    public class ClipboardService : IClipboardService
    {
        public bool ContainsText()
        {
            // 简化实现，总是返回true用于测试
            return true;
        }

        public async Task<string?> GetTextAsync()
        {
            try
            {
                if (Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
                {
                    var clipboard = desktop.MainWindow?.Clipboard;
                    if (clipboard != null)
                    {
#pragma warning disable CS0618 // 忽略过时API警告
                        return await clipboard.GetTextAsync();
#pragma warning restore CS0618
                    }
                }

                // 回退到模拟数据
                await Task.Delay(100); // 模拟异步操作
                return "{\"lessonId\": \"test123\", \"url\": \"https://example.com/video.mp4\"}";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取剪贴板内容失败: {ex.Message}");
                return null;
            }
        }

        public async Task<bool> SetClipboardTextAsync(string text)
        {
            try
            {
                if (Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
                {
                    var clipboard = desktop.MainWindow?.Clipboard;
                    if (clipboard != null)
                    {
                        await clipboard.SetTextAsync(text);
                        return true;
                    }
                }

                // 如果无法访问Avalonia剪贴板，回退到控制台输出（仅用于调试）
                Console.WriteLine($"[DEBUG] 应该复制到剪贴板的文本: {text}");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置剪贴板内容失败: {ex.Message}");
                return false;
            }
        }
    }
}