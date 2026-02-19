using System.ComponentModel;
using Classin视频解析下载工具.Shared.Commands;
using Classin视频解析下载工具.Shared.Constants;
using Classin视频解析下载工具.Models;
using Classin视频解析下载工具.Services.CoreServices;

namespace Classin视频解析下载工具.ViewModels
{
    public class SettingsViewModel : ViewModelBase
    {
        private readonly IConfigurationManager _configurationManager;
        private DownloadSettings _settings;

        public SettingsViewModel(IConfigurationManager configurationManager)
        {
            _configurationManager = configurationManager;
            _settings = _configurationManager.GetConfiguration();
        }

        public DownloadSettings Settings
        {
            get => _settings;
            set => SetProperty(ref _settings, value);
        }

        public int MaxConcurrentDownloads
        {
            get => _settings.MaxConcurrentDownloads;
            set
            {
                if (_settings.MaxConcurrentDownloads != value)
                {
                    _settings.MaxConcurrentDownloads = value;
                    OnPropertyChanged();
                    SaveSettings();
                }
            }
        }



        public string DownloadPath
        {
            get => _settings.DownloadPath;
            set
            {
                if (_settings.DownloadPath != value)
                {
                    _settings.DownloadPath = value;
                    OnPropertyChanged();
                    SaveSettings();
                }
            }
        }

        public bool EnableLogging
        {
            get => _settings.EnableLogging;
            set
            {
                if (_settings.EnableLogging != value)
                {
                    _settings.EnableLogging = value;
                    OnPropertyChanged();
                    SaveSettings();
                }
            }
        }

        public bool AutoCheckUpdates
        {
            get => _settings.AutoCheckUpdates;
            set
            {
                if (_settings.AutoCheckUpdates != value)
                {
                    _settings.AutoCheckUpdates = value;
                    OnPropertyChanged();
                    SaveSettings();
                }
            }
        }

        public int BufferSizeKB
        {
            get => _settings.BufferSizeKB;
            set
            {
                if (_settings.BufferSizeKB != value)
                {
                    _settings.BufferSizeKB = value;
                    OnPropertyChanged();
                    SaveSettings();
                }
            }
        }

        public double TimeoutHours
        {
            get => _settings.TimeoutHours;
            set
            {
                if (_settings.TimeoutHours != value)
                {
                    _settings.TimeoutHours = value;
                    OnPropertyChanged();
                    SaveSettings();
                }
            }
        }

        public int MaxRetries
        {
            get => _settings.MaxRetries;
            set
            {
                if (_settings.MaxRetries != value)
                {
                    _settings.MaxRetries = value;
                    OnPropertyChanged();
                    SaveSettings();
                }
            }
        }

        public string DefaultUserAgent
        {
            get => _settings.DefaultUserAgent;
            set
            {
                if (_settings.DefaultUserAgent != value)
                {
                    _settings.DefaultUserAgent = value;
                    OnPropertyChanged();
                    SaveSettings();
                }
            }
        }

        public string DefaultReferrer
        {
            get => _settings.DefaultReferrer;
            set
            {
                if (_settings.DefaultReferrer != value)
                {
                    _settings.DefaultReferrer = value;
                    OnPropertyChanged();
                    SaveSettings();
                }
            }
        }

        private void LoadSettings()
        {
            _settings = _configurationManager.GetConfiguration();
        }

        private void SaveSettings()
        {
            _configurationManager.UpdateConfiguration(c =>
            {
                c.MaxConcurrentDownloads = _settings.MaxConcurrentDownloads;
                c.DownloadPath = _settings.DownloadPath;
                c.BufferSizeKB = _settings.BufferSizeKB;
                c.TimeoutHours = _settings.TimeoutHours;
                c.MaxRetries = _settings.MaxRetries;
                c.EnableLogging = _settings.EnableLogging;
                c.AutoCheckUpdates = _settings.AutoCheckUpdates;
                c.DefaultUserAgent = _settings.DefaultUserAgent;
                c.DefaultReferrer = _settings.DefaultReferrer;
            });
        }

        public void ResetToDefaults()
        {
            _settings = new DownloadSettings
            {
                MaxConcurrentDownloads = AppConstants.MaxConcurrentDownloads,
                DownloadPath = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, AppConstants.DefaultDownloadFolder),
                BufferSizeKB = AppConstants.DefaultBufferSize / 1024,
                TimeoutHours = AppConstants.DefaultTimeoutHours,
                MaxRetries = AppConstants.DefaultMaxRetries,
                EnableLogging = true,
                AutoCheckUpdates = true,
                DefaultUserAgent = AppConstants.DefaultUserAgent,
                DefaultReferrer = AppConstants.DefaultReferrer
            };

            OnPropertyChanged(nameof(Settings));
            OnPropertyChanged(nameof(MaxConcurrentDownloads));
            OnPropertyChanged(nameof(DownloadPath));
            OnPropertyChanged(nameof(BufferSizeKB));
            OnPropertyChanged(nameof(TimeoutHours));
            OnPropertyChanged(nameof(MaxRetries));
            OnPropertyChanged(nameof(EnableLogging));
            OnPropertyChanged(nameof(AutoCheckUpdates));
            OnPropertyChanged(nameof(DefaultUserAgent));
            OnPropertyChanged(nameof(DefaultReferrer));

            SaveSettings();
        }
    }
}
