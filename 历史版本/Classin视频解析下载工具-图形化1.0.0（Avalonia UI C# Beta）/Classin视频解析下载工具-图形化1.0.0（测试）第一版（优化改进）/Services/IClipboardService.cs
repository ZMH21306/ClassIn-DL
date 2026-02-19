using System.Threading.Tasks;

namespace Classin视频解析下载工具.Services
{
    public interface IClipboardService
    {
        Task<string?> GetTextAsync();
        Task<bool> SetClipboardTextAsync(string text);
        bool ContainsText();
    }
}
