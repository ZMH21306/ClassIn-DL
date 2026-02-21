using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Classin视频解析下载工具.Models;

namespace Classin视频解析下载工具.Services.DownloadServices
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
}