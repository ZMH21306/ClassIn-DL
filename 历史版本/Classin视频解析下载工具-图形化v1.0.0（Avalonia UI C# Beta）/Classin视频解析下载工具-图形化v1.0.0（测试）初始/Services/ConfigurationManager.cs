using System;
using System.IO;

namespace VideoDownloader.Services
{
    public interface IConfigurationManager
    {
        Models.DownloadSettings GetConfiguration();
        void SaveConfiguration(Models.DownloadSettings settings);
        void UpdateConfiguration(Action<Models.DownloadSettings> updateAction);
        string GetConfigurationFilePath();
    }

    public class ConfigurationManager : IConfigurationManager
    {
        private readonly string _configFilePath;
        private Models.DownloadSettings? _cachedSettings;

        public ConfigurationManager()
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var appFolder = Path.Combine(appDataPath, "Classin视频解析下载工具");
            Directory.CreateDirectory(appFolder);
            _configFilePath = Path.Combine(appFolder, "appsettings.json");
        }

        public Models.DownloadSettings GetConfiguration()
        {
            if (_cachedSettings != null)
                return _cachedSettings;

            try
            {
                if (File.Exists(_configFilePath))
                {
                    var json = File.ReadAllText(_configFilePath);
                    _cachedSettings = System.Text.Json.JsonSerializer.Deserialize<Models.DownloadSettings>(json);
                }
            }
            catch
            {
                // 忽略序列化错误
            }

            _cachedSettings ??= new Models.DownloadSettings
            {
                MaxConcurrentDownloads = Constants.AppConstants.MaxConcurrentDownloads,
                MaxDownloadThreads = Constants.AppConstants.MaxDownloadThreads,
                DownloadPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, VideoDownloader.Constants.AppConstants.DefaultDownloadFolder),
                BufferSizeKB = VideoDownloader.Constants.AppConstants.DefaultBufferSize / 1024,
                TimeoutHours = VideoDownloader.Constants.AppConstants.DefaultTimeoutHours,
                MaxRetries = VideoDownloader.Constants.AppConstants.DefaultMaxRetries,
                EnableLogging = true,
                AutoCheckUpdates = true,
                DefaultUserAgent = VideoDownloader.Constants.AppConstants.DefaultUserAgent,
                DefaultReferrer = VideoDownloader.Constants.AppConstants.DefaultReferrer
            };

            return _cachedSettings;
        }

        public void SaveConfiguration(Models.DownloadSettings settings)
        {
            try
            {
                var json = System.Text.Json.JsonSerializer.Serialize(settings, new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(_configFilePath, json);
                _cachedSettings = settings;
            }
            catch
            {
                // 忽略保存错误
            }
        }

        public void UpdateConfiguration(Action<Models.DownloadSettings> updateAction)
        {
            var settings = GetConfiguration();
            updateAction(settings);
            SaveConfiguration(settings);
        }

        public string GetConfigurationFilePath()
        {
            return _configFilePath;
        }
    }
}