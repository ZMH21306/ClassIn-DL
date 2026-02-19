using System;
using System.IO;
using System.Collections.Generic;
using Classin视频解析下载工具.Models;
using Classin视频解析下载工具.Constants;

namespace Classin视频解析下载工具.Services
{
    public interface IConfigurationManager
    {
        DownloadSettings GetConfiguration();
        void SaveConfiguration(DownloadSettings settings);
        void UpdateConfiguration(Action<DownloadSettings> updateAction);
        string GetConfigurationFilePath();
    }

    public class ConfigurationManager : IConfigurationManager
    {
        private readonly string _configFilePath;
        private DownloadSettings? _cachedSettings;

        public ConfigurationManager()
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var appFolder = Path.Combine(appDataPath, "Classin视频解析下载工具");
            Directory.CreateDirectory(appFolder);
            _configFilePath = Path.Combine(appFolder, "appsettings.json");
        }

        public DownloadSettings GetConfiguration()
        {
            DownloadSettings settings;

            try
            {
                if (File.Exists(_configFilePath))
                {
                    var json = File.ReadAllText(_configFilePath);
                    settings = System.Text.Json.JsonSerializer.Deserialize<DownloadSettings>(json) ?? new DownloadSettings();
                    Console.WriteLine($"从配置文件读取下载路径: {settings.DownloadPath}");
                }
                else
                {
                    settings = new DownloadSettings();
                    Console.WriteLine("配置文件不存在，使用默认配置");
                }
            }
            catch (Exception ex)
            {
                // 忽略序列化错误，使用默认配置
                settings = new DownloadSettings();
                Console.WriteLine($"读取配置文件失败: {ex.Message}");
            }

            // 设置默认值
            settings.MaxConcurrentDownloads = settings.MaxConcurrentDownloads > 0 ? settings.MaxConcurrentDownloads : AppConstants.MaxConcurrentDownloads;
            settings.MaxDownloadThreads = settings.MaxDownloadThreads > 0 ? settings.MaxDownloadThreads : AppConstants.MaxDownloadThreads;
            settings.BufferSizeKB = settings.BufferSizeKB > 0 ? settings.BufferSizeKB : AppConstants.DefaultBufferSize / 1024;
            settings.TimeoutHours = settings.TimeoutHours > 0 ? settings.TimeoutHours : AppConstants.DefaultTimeoutHours;
            settings.MaxRetries = settings.MaxRetries > 0 ? settings.MaxRetries : AppConstants.DefaultMaxRetries;
            settings.EnableLogging = settings.EnableLogging;
            settings.AutoCheckUpdates = settings.AutoCheckUpdates;
            settings.DefaultUserAgent = !string.IsNullOrEmpty(settings.DefaultUserAgent) ? settings.DefaultUserAgent : AppConstants.DefaultUserAgent;
            settings.DefaultReferrer = !string.IsNullOrEmpty(settings.DefaultReferrer) ? settings.DefaultReferrer : AppConstants.DefaultReferrer;
            settings.CustomSettings ??= new Dictionary<string, object>();

            // 强制使用软件安装目录下的"下载目录"文件夹作为默认下载路径
            string newDownloadPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, AppConstants.DefaultDownloadFolder);
            Console.WriteLine($"设置新的下载路径: {newDownloadPath}");
            settings.DownloadPath = newDownloadPath;

            // 确保下载目录存在
            try
            {
                Directory.CreateDirectory(settings.DownloadPath);
                Console.WriteLine($"创建下载目录成功: {settings.DownloadPath}");
            }
            catch (Exception ex)
            {
                // 忽略目录创建错误，避免权限问题导致程序启动失败
                Console.WriteLine($"创建下载目录失败: {ex.Message}");
            }

            // 保存更新后的配置
            try
            {
                Console.WriteLine($"保存配置到: {_configFilePath}");
                SaveConfiguration(settings);
                Console.WriteLine("保存配置成功");
            }
            catch (Exception ex)
            {
                // 忽略保存错误
                Console.WriteLine($"保存配置失败: {ex.Message}");
            }

            // 更新缓存
            _cachedSettings = settings;

            return settings;
        }

        public void SaveConfiguration(DownloadSettings settings)
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

        public void UpdateConfiguration(Action<DownloadSettings> updateAction)
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