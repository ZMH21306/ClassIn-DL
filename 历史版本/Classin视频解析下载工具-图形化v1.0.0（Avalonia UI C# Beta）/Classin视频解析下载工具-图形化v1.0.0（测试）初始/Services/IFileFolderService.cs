using System.Threading.Tasks;

namespace VideoDownloader.Services
{
    public interface IFileFolderService
    {
        Task<string?> SelectFolderAsync(string title, string initialDirectory);
        Task<string?> SelectFileAsync(string title, string initialDirectory, string filter);
        bool CreateDirectory(string path);
        bool DirectoryExists(string path);
        bool FileExists(string path);
        void OpenDirectory(string path);
        void DeleteFile(string path);
        void DeleteDirectory(string path, bool recursive = false);
    }
}
