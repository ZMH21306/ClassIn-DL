using System;
using System.Threading.Tasks;

namespace Classin视频解析下载工具.Services
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
                // 简化实现，返回模拟的剪贴板内容
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
                // 简化实现，总是返回成功
                await Task.Delay(50); // 模拟异步操作
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置剪贴板内容失败: {ex.Message}");
                return false;
            }
        }
    }
}