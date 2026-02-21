using System.Threading.Tasks;
using Avalonia.Controls;

namespace Classin视频解析下载工具.Services.SystemServices
{
    public interface IFileFolderService
    {
        void SetMainWindow(Window mainWindow);
        Task<string?> SelectFolderAsync(string title, string initialDirectory);
        Task<string?> SelectFileAsync(string title, string initialDirectory, string filter);
        bool CreateDirectory(string path);
        bool DirectoryExists(string path);
        bool FileExists(string path);
        Task OpenDirectoryAsync(string path);
        void OpenDirectory(string path);
        void DeleteFile(string path);
        void DeleteDirectory(string path, bool recursive = false);
    }
}
