using System;
using System.IO;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Platform.Storage;

namespace Classin视频解析下载工具.Services.SystemServices
{
    public class FileFolderService : IFileFolderService
    {
        private Window? _mainWindow;

        public FileFolderService(Window? mainWindow = null)
        {
            _mainWindow = mainWindow;
        }

        public void SetMainWindow(Window mainWindow)
        {
            _mainWindow = mainWindow;
        }

        public async Task<string?> SelectFolderAsync(string title, string initialDirectory = "")
        {
            if (_mainWindow == null) return null;

            try
            {
                var folders = await _mainWindow.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
                {
                    Title = title,
                    AllowMultiple = false
                });

                return folders.Count > 0 ? folders[0].TryGetLocalPath() : null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"选择文件夹失败: {ex.Message}");
                return null;
            }
        }

        public async Task<string?> SelectFileAsync(string title, string initialDirectory, string filter)
        {
            if (_mainWindow == null) return null;

            try
            {
                var files = await _mainWindow.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
                {
                    Title = title,
                    AllowMultiple = false
                });

                return files.Count > 0 ? files[0].TryGetLocalPath() : null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"选择文件失败: {ex.Message}");
                return null;
            }
        }

        public bool CreateDirectory(string path)
        {
            try
            {
                Directory.CreateDirectory(path);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建目录失败: {ex.Message}");
                return false;
            }
        }

        public bool DirectoryExists(string path)
        {
            return Directory.Exists(path);
        }

        public bool FileExists(string path)
        {
            return File.Exists(path);
        }

        public async Task OpenDirectoryAsync(string path)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    await System.Threading.Tasks.Task.Run(() =>
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = "explorer.exe",
                            Arguments = path,
                            UseShellExecute = true,
                            CreateNoWindow = true
                        });
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打开目录失败: {ex.Message}");
            }
        }

        // 保持原有同步方法以兼容性
        public void OpenDirectory(string path)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = path,
                        UseShellExecute = true,
                        CreateNoWindow = true
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打开目录失败: {ex.Message}");
            }
        }

        public void DeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"删除文件失败: {ex.Message}");
            }
        }

        public void DeleteDirectory(string path, bool recursive = false)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    Directory.Delete(path, recursive);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"删除目录失败: {ex.Message}");
            }
        }

        // 额外的方法，用于兼容原有代码
        public string GetDirectoryName(string path)
        {
            return Path.GetDirectoryName(path) ?? string.Empty;
        }
    }
}