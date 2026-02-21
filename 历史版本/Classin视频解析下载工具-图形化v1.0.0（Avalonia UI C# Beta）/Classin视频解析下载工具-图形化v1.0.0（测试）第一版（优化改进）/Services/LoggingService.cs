using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Classin视频解析下载工具.Services
{
    public interface ILoggingService : IDisposable
    {
        void Info(string message, string category = "");
        void Warning(string message, string category = "");
        void Error(string message, string category = "");
        void Debug(string message, string category = "");
        event EventHandler<string>? LogUpdated;
    }

    public class LoggingService : ILoggingService
    {
        private bool _disposed;
        private readonly object _lockObject = new();

        public event EventHandler<string>? LogUpdated;

        public void Info(string message, string category = "")
        {
            LogMessage("INFO", message, category);
        }

        public void Warning(string message, string category = "")
        {
            LogMessage("WARN", message, category);
        }

        public void Error(string message, string category = "")
        {
            LogMessage("ERROR", message, category);
        }

        public void Debug(string message, string category = "")
        {
            LogMessage("DEBUG", message, category);
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
                    LogUpdated?.Invoke(this, formattedMessage);
                    System.Diagnostics.Debug.WriteLine(formattedMessage);
                }
            }
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