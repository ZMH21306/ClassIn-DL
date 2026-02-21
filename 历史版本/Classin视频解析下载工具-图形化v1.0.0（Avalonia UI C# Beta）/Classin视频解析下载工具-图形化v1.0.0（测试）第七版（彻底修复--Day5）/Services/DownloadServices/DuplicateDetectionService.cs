/*
 * 模块名称：DuplicateDetectionService
 * 功能描述：提供视频文件重复检测服务，使用LRU缓存策略管理文件名称缓存
 * 主要依赖：
 *   - System.Collections.Generic (HashSet, Dictionary, LinkedList)
 *   - System.IO (文件操作)
 *   - Classin视频解析下载工具.Models (VideoItem模型)
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Classin视频解析下载工具.Models;
using Classin视频解析下载工具.Services.CoreServices;

namespace Classin视频解析下载工具.Services.DownloadServices
{
    public class DuplicateDetectionService : IDuplicateDetectionService
    {
        private readonly HashSet<string> _fileNameCache = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _nameToPathMap = new(StringComparer.OrdinalIgnoreCase);
        private readonly LinkedList<string> _lruList = new(); // 用于实现LRU缓存策略
        private readonly object _cacheLock = new();
        private readonly ILoggingService _loggingService;
        private string _downloadPath = string.Empty;
        private const int MAX_CACHE_SIZE = 1000; // 缓存大小限制

        public DuplicateDetectionService(ILoggingService loggingService)
        {
            _loggingService = loggingService ?? throw new ArgumentNullException(nameof(loggingService));
        }

        /// <summary>
        /// 设置下载路径并刷新缓存
        /// </summary>
        /// <param name="downloadPath">下载路径</param>
        /// <exception cref="ArgumentNullException">当下载路径为 null 时抛出</exception>
        public void SetDownloadPath(string downloadPath)
        {
            if (downloadPath == null)
                throw new ArgumentNullException(nameof(downloadPath), "下载路径不能为空");

            _downloadPath = downloadPath;
            _loggingService.Info($"下载路径已设置为: {downloadPath}", "DuplicateDetection");
            RefreshCache();
        }

        /// <summary>
        /// 检查文件是否重复
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <returns>如果文件已存在于缓存中返回 true，否则返回 false</returns>
        /// <exception cref="ArgumentException">当文件名称为空时抛出</exception>
        public bool IsDuplicate(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                throw new ArgumentException("文件名称不能为空", nameof(fileName));

            lock (_cacheLock)
            {
                bool isDuplicate = _fileNameCache.Contains(fileName);

                // 如果找到，更新LRU列表
                if (isDuplicate)
                {
                    UpdateLRU(fileName);
                }

                return isDuplicate;
            }
        }

        /// <summary>
        /// 将视频项添加到缓存
        /// </summary>
        /// <param name="item">视频项</param>
        /// <exception cref="ArgumentNullException">当视频项为 null 时抛出</exception>
        /// <exception cref="ArgumentException">当视频名称为空时抛出</exception>
        public void AddToCache(VideoItem item)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item), "视频项不能为空");
            if (item.Name == null)
                throw new ArgumentException("视频名称不能为空", nameof(item));

            var safeName = SanitizeFileName(item.Name);
            AddToCacheInternal(safeName);
        }

        /// <summary>
        /// 将文件名称添加到缓存
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <exception cref="ArgumentException">当文件名称为空时抛出</exception>
        public void AddToCache(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                throw new ArgumentException("文件名称不能为空", nameof(fileName));

            var safeName = SanitizeFileName(fileName);
            AddToCacheInternal(safeName);
        }

        private void AddToCacheInternal(string safeName)
        {
            lock (_cacheLock)
            {
                try
                {
                    // 检查缓存大小，如果超过限制，清理最不常用的项
                    if (_fileNameCache.Count >= MAX_CACHE_SIZE && !_fileNameCache.Contains(safeName))
                    {
                        CleanupCache();
                    }

                    // 添加到缓存
                    _fileNameCache.Add(safeName);

                    var fullPath = Path.Combine(_downloadPath, $"{safeName}.mp4");
                    _nameToPathMap[safeName] = fullPath;

                    // 更新LRU列表
                    UpdateLRU(safeName);

                    _loggingService.Debug($"文件已添加到缓存: {safeName}", "DuplicateDetection");
                }
                catch (Exception ex)
                {
                    _loggingService.Error($"添加文件到缓存失败: {ex.Message}", "DuplicateDetection");
                }
            }
        }

        /// <summary>
        /// 从缓存中移除文件名称
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <exception cref="ArgumentException">当文件名称为空时抛出</exception>
        public void RemoveFromCache(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                throw new ArgumentException("文件名称不能为空", nameof(fileName));

            lock (_cacheLock)
            {
                try
                {
                    _fileNameCache.Remove(fileName);
                    _nameToPathMap.Remove(fileName);
                    _lruList.Remove(fileName);
                    _loggingService.Debug($"文件已从缓存移除: {fileName}", "DuplicateDetection");
                }
                catch (Exception ex)
                {
                    _loggingService.Error($"从缓存移除文件失败: {ex.Message}", "DuplicateDetection");
                }
            }
        }

        /// <summary>
        /// 清空所有缓存
        /// </summary>
        public void ClearCache()
        {
            lock (_cacheLock)
            {
                try
                {
                    _fileNameCache.Clear();
                    _nameToPathMap.Clear();
                    _lruList.Clear();
                    _loggingService.Info("缓存已清空", "DuplicateDetection");
                }
                catch (Exception ex)
                {
                    _loggingService.Error($"清空缓存失败: {ex.Message}", "DuplicateDetection");
                }
            }
        }

        /// <summary>
        /// 清理文件名，移除无效字符并限制长度
        /// </summary>
        /// <param name="name">原始文件名</param>
        /// <returns>清理后的安全文件名</returns>
        /// <exception cref="ArgumentException">当文件名为空时抛出</exception>
        public string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("文件名称不能为空", nameof(name));

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
            lock (_cacheLock)
            {
                _fileNameCache.Clear();
                _nameToPathMap.Clear();
                _lruList.Clear();

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
                        _lruList.AddLast(fileName);
                    }
                    _loggingService.Debug($"缓存刷新完成，加载了 {mp4Files.Length} 个文件", "DuplicateDetection");
                }
                catch (Exception ex)
                {
                    _loggingService.Error($"刷新缓存失败: {ex.Message}", "DuplicateDetection");
                }
            }
        }

        /// <summary>
        /// 更新LRU列表，将项移到末尾（表示最近使用）
        /// </summary>
        /// <param name="fileName">文件名</param>
        private void UpdateLRU(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return;

            _lruList.Remove(fileName);
            _lruList.AddLast(fileName);
        }

        /// <summary>
        /// 清理缓存，移除最不常用的项
        /// </summary>
        private void CleanupCache()
        {
            // 清理20%的缓存项
            int itemsToRemove = Math.Max(1, MAX_CACHE_SIZE / 5);

            for (int i = 0; i < itemsToRemove && _lruList.Count > 0; i++)
            {
                if (_lruList.First != null)
                {
                    string leastUsed = _lruList.First.Value;
                    _lruList.RemoveFirst();
                    _fileNameCache.Remove(leastUsed);
                    _nameToPathMap.Remove(leastUsed);
                }
            }
        }
    }
}