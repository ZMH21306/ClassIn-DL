using System;
using System.Collections.Generic;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;

namespace VideoDownloader.ViewModels
{
    public class LogViewModel : ViewModelBase, IDisposable
    {
        private readonly StringBuilder _logBuffer;
        private readonly object _lockObject = new();
        private bool _isInitialized;
        private bool _disposed;

        public LogViewModel()
        {
            _logBuffer = new StringBuilder();
        }

        private string _logText = string.Empty;
        public string LogText
        {
            get => _logText;
            private set => SetProperty(ref _logText, value);
        }

        public bool IsInitialized
        {
            get => _isInitialized;
            set => SetProperty(ref _isInitialized, value);
        }

        public void AppendLog(string message)
        {
            if (!_isInitialized || _disposed) return;

            lock (_lockObject)
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                var logEntry = $"[{timestamp}] {message}";
                
                _logBuffer.AppendLine(logEntry);
                
                // 限制日志缓冲区大小
                if (_logBuffer.Length > 100000) // 100KB
                {
                    var content = _logBuffer.ToString();
                    var lines = content.Split('\n');
                    if (lines.Length > 1000)
                    {
                        _logBuffer.Clear();
                        var startIndex = Math.Max(0, lines.Length - 500);
                        _logBuffer.Append(string.Join("\n", lines[startIndex..]));
                        _logBuffer.AppendLine("--- 日志已截断 ---");
                    }
                }

                LogText = _logBuffer.ToString();
            }
        }

        public void AppendLogForSlider(string message)
        {
            AppendLog(message);
        }

        public void ClearLog()
        {
            lock (_lockObject)
            {
                _logBuffer.Clear();
                LogText = string.Empty;
            }
        }

        public string GetLogContent()
        {
            lock (_lockObject)
            {
                return _logBuffer.ToString();
            }
        }

        public void SaveLogToFile(string filePath)
        {
            try
            {
                lock (_lockObject)
                {
                    System.IO.File.WriteAllText(filePath, _logBuffer.ToString(), Encoding.UTF8);
                }
            }
            catch
            {
                // 忽略保存错误
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            
            // 保存日志到文件
            try
            {
                var logDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                                        "Classin视频解析下载工具", "Logs");
                System.IO.Directory.CreateDirectory(logDir);
                
                var logFile = System.IO.Path.Combine(logDir, $"app_{DateTime.Now:yyyyMMdd}.log");
                SaveLogToFile(logFile);
            }
            catch
            {
                // 忽略保存错误
            }
        }
    }
}