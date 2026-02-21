using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Classin视频解析下载工具.Models;

namespace Classin视频解析下载工具.Services.CoreServices
{
    public interface IConfigurationManager
    {
        DownloadSettings GetConfiguration();
        void SaveConfiguration(DownloadSettings settings);
        void UpdateConfiguration(Action<DownloadSettings> updateAction);
        string GetConfigurationFilePath();
        Task<DownloadSettings> GetConfigurationAsync();
        Task SaveConfigurationAsync(DownloadSettings settings);
    }
}