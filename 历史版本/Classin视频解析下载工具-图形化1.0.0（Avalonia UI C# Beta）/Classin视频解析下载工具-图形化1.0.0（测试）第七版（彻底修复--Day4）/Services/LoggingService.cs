using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Classin视频解析下载工具.Services
{
    public class LoggingService : ILoggingService
    {
        private bool _disposed;
        private readonly object _lockObject = new();
        private LogLevel _minLogLevel = LogLevel.DEBUG; // 默认记录DEBUG及以上级别，确保所有操作都有日志
        private static readonly Dictionary<LogLevel, string> LevelNames = new()
        {
            { LogLevel.ALL, "ALL" },
            { LogLevel.TRACE, "TRACE" },
            { LogLevel.DEBUG, "DEBUG" },
            { LogLevel.INFO, "INFO" },
            { LogLevel.WARN, "WARN" },
            { LogLevel.ERROR, "ERROR" },
            { LogLevel.FATAL, "FATAL" },
            { LogLevel.OFF, "OFF" }
        };

        public event EventHandler<string>? LogUpdated;

        public void Info(string message, string category = "")
        {
            Log(LogLevel.INFO, message, category);
        }

        public void Warning(string message, string category = "")
        {
            Log(LogLevel.WARN, message, category);
        }

        public void Error(string message, string category = "")
        {
            Log(LogLevel.ERROR, message, category);
        }

        public void Debug(string message, string category = "")
        {
            Log(LogLevel.DEBUG, message, category);
        }

        public void Trace(string message, string category = "")
        {
            Log(LogLevel.TRACE, message, category);
        }

        public void Fatal(string message, string category = "")
        {
            Log(LogLevel.FATAL, message, category);
        }

        public void Log(LogLevel level, string message, string category = "")
        {
            // 检查日志级别是否启用
            if (level < _minLogLevel || level == LogLevel.OFF)
                return;

            var levelName = LevelNames.TryGetValue(level, out var name) ? name : "UNKNOWN";
            LogMessage(levelName, message, category);
        }

        private void LogMessage(string level, string message, string category)
        {
            if (_disposed) return;

            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var formattedMessage = string.IsNullOrEmpty(category)
                ? $"[{timestamp}] [{level}] {message}"
                : $"[{timestamp}] [{level}] [{category}] {message}";
            
            lock (_lockObject)
            {
                if (!_disposed)
                {
                    try
                    {
                        LogUpdated?.Invoke(this, formattedMessage);
                        System.Diagnostics.Debug.WriteLine(formattedMessage);
                    }
                    catch (Exception ex)
                    {
                        // 日志记录本身不应该抛出异常
                        System.Diagnostics.Debug.WriteLine($"[ERROR] 日志记录失败: {ex.Message}");
                    }
                }
            }
        }

        public void SetLogLevel(LogLevel minLevel)
        {
            _minLogLevel = minLevel;
            Log(LogLevel.INFO, $"日志级别已设置为: {LevelNames[minLevel]}", "LoggingService");
        }

        public LogLevel GetCurrentLogLevel()
        {
            return _minLogLevel;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            
            lock (_lockObject)
            {
                LogUpdated = null;
            }
        }
    }
}