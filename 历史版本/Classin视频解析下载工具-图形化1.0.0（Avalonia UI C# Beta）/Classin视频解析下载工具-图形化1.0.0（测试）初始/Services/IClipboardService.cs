using System.Threading.Tasks;

namespace VideoDownloader.Services
{
    public interface IClipboardService
    {
        Task<string?> GetTextAsync();
        Task<bool> SetClipboardTextAsync(string text);
        bool ContainsText();
    }
}
