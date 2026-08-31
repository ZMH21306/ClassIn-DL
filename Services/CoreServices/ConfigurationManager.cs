using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Classin视频解析下载工具.Models;
using Classin视频解析下载工具.Shared.Constants;

namespace Classin视频解析下载工具.Services.CoreServices
{
    public class ConfigurationManager : IConfigurationManager
    {
        private readonly string _configFilePath;
        private Lazy<DownloadSettings> _cachedSettings;

        public ConfigurationManager()
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var appFolder = Path.Combine(appDataPath, "Classin视频解析下载工具");
            // 延迟创建目录，只在需要时创建
            _configFilePath = Path.Combine(appFolder, "appsettings.json");

            // 使用Lazy<T>实现延迟加载配置
            _cachedSettings = new Lazy<DownloadSettings>(() => LoadConfiguration());
        }

        public DownloadSettings GetConfiguration()
        {
            return _cachedSettings.Value;
        }

        private DownloadSettings LoadConfiguration()
        {
            DownloadSettings settings;

            try
            {
                if (File.Exists(_configFilePath))
                {
                    var json = File.ReadAllText(_configFilePath);
                    var options = new System.Text.Json.JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    };
                    settings = System.Text.Json.JsonSerializer.Deserialize<DownloadSettings>(json, options) ?? new DownloadSettings();
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
                var appFolder = Path.GetDirectoryName(_configFilePath);
                if (!string.IsNullOrEmpty(appFolder))
                {
                    Directory.CreateDirectory(appFolder);
                }
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

            return settings;
        }

        public void SaveConfiguration(DownloadSettings settings)
        {
            // 确保 PascalCase 属性名
            var options = new System.Text.Json.JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = null // 保持 PascalCase
            };
            var json = System.Text.Json.JsonSerializer.Serialize(settings, options);

            var appFolder = Path.GetDirectoryName(_configFilePath);
            if (!string.IsNullOrEmpty(appFolder) && !Directory.Exists(appFolder))
            {
                Directory.CreateDirectory(appFolder);
            }

            // 重试机制（最多 3 次）
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    File.WriteAllText(_configFilePath, json);
                    _cachedSettings = new Lazy<DownloadSettings>(() => settings);
                    return; // 保存成功，退出重试循环
                }
                catch (IOException ex) when (attempt < 3)
                {
                    // 文件可能被其他进程占用，等待后重试
                    System.Diagnostics.Debug.WriteLine($"保存配置失败（尝试 {attempt}/3）: {ex.Message}");
                    Thread.Sleep(200 * attempt); // 指数退避
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"保存配置失败（尝试 {attempt}/3）: {ex.Message}");
                    if (attempt == 3)
                    {
                        // 最终尝试失败，记录日志
                        System.Diagnostics.Debug.WriteLine($"配置保存最终失败，建议检查磁盘空间和文件权限");
                    }
                }
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

        // 异步版本的获取配置方法
        public async Task<DownloadSettings> GetConfigurationAsync()
        {
            if (_cachedSettings.IsValueCreated)
            {
                return _cachedSettings.Value;
            }

            return await Task.Run(() => GetConfiguration());
        }

        // 异步版本的保存配置方法
        public async Task SaveConfigurationAsync(DownloadSettings settings)
        {
            await Task.Run(() => SaveConfiguration(settings));
        }
    }
}