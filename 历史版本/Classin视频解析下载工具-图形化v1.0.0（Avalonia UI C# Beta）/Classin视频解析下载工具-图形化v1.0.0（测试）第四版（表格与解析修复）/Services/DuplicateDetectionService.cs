using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Classin视频解析下载工具.Models;

namespace Classin视频解析下载工具.Services
{
    public interface IDuplicateDetectionService
    {
        void SetDownloadPath(string downloadPath);
        bool IsDuplicate(string fileName);
        void AddToCache(VideoItem item);
        void AddToCache(string fileName);
        void RemoveFromCache(string fileName);
        void ClearCache();
        string SanitizeFileName(string name);
    }

    public class DuplicateDetectionService : IDuplicateDetectionService
    {
        private readonly HashSet<string> _fileNameCache = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _nameToPathMap = new(StringComparer.OrdinalIgnoreCase);
        private string _downloadPath = string.Empty;

        public void SetDownloadPath(string downloadPath)
        {
            _downloadPath = downloadPath;
            RefreshCache();
        }

        public bool IsDuplicate(string fileName)
        {
            return _fileNameCache.Contains(fileName);
        }

        public void AddToCache(VideoItem item)
        {
            var safeName = SanitizeFileName(item.Name);
            _fileNameCache.Add(safeName);
            
            var fullPath = Path.Combine(_downloadPath, $"{safeName}.mp4");
            _nameToPathMap[safeName] = fullPath;
        }

        public void AddToCache(string fileName)
        {
            var safeName = SanitizeFileName(fileName);
            _fileNameCache.Add(safeName);
            
            var fullPath = Path.Combine(_downloadPath, $"{safeName}.mp4");
            _nameToPathMap[safeName] = fullPath;
        }

        public void RemoveFromCache(string fileName)
        {
            _fileNameCache.Remove(fileName);
            _nameToPathMap.Remove(fileName);
        }

        public void ClearCache()
        {
            _fileNameCache.Clear();
            _nameToPathMap.Clear();
        }

        public string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return name;

            var invalidChars = Path.GetInvalidFileNameChars();
            var builder = new System.Text.StringBuilder();

            foreach (var c in name)
            {
                builder.Append(invalidChars.Contains(c) ? '_' : c);
            }

            var result = builder.ToString();
            
            // 限制文件名长度
            if (result.Length > 150)
            {
                result = result.Substring(0, 150);
            }

            return result;
        }

        private void RefreshCache()
        {
            _fileNameCache.Clear();
            _nameToPathMap.Clear();

            if (string.IsNullOrEmpty(_downloadPath) || !Directory.Exists(_downloadPath))
                return;

            try
            {
                var mp4Files = Directory.GetFiles(_downloadPath, "*.mp4", SearchOption.TopDirectoryOnly);
                foreach (var file in mp4Files)
                {
                    var fileName = Path.GetFileNameWithoutExtension(file);
                    _fileNameCache.Add(fileName);
                    _nameToPathMap[fileName] = file;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"刷新缓存失败: {ex.Message}");
            }
        }
    }
}