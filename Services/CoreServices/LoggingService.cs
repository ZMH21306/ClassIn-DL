using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Classin视频解析下载工具.Services.CoreServices
{
    public class LoggingService : ILoggingService
    {
        private bool _disposed;
        private readonly object _lockObject = new();
        private LogLevel _minLogLevel = LogLevel.DEBUG;
        private static readonly Dictionary<LogLevel, string> LevelNames = new()
        {
            { LogLevel.ALL, "ALL" }, { LogLevel.TRACE, "TRACE" },
            { LogLevel.DEBUG, "DEBUG" }, { LogLevel.INFO, "INFO" },
            { LogLevel.WARN, "WARN" }, { LogLevel.ERROR, "ERROR" },
            { LogLevel.FATAL, "FATAL" }, { LogLevel.OFF, "OFF" }
        };

        private readonly string _logFilePath;
        private readonly object _fileLock = new();
        private readonly ConcurrentQueue<string> _pendingWrites = new();
        private readonly SemaphoreSlim _writeSignal = new(0, int.MaxValue);
        private readonly CancellationTokenSource _cts = new();
        private readonly Task _fileWriterTask;
        private long _currentLogSize;
        private const long MaxLogFileSize = 5 * 1024 * 1024;
        private const int MaxBackupFiles = 5;
        private const int BufferFlushIntervalMs = 500;

        public event EventHandler<string>? LogUpdated;

        public LoggingService(string? customLogPath = null)
        {
            if (!string.IsNullOrEmpty(customLogPath))
            {
                _logFilePath = customLogPath;
            }
            else
            {
                try
                {
                    var logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
                    Directory.CreateDirectory(logDir);
                    _logFilePath = Path.Combine(logDir, "app.log");
                }
                catch
                {
                    var tempDir = Path.GetTempPath();
                    _logFilePath = Path.Combine(tempDir, "ClassInVideoDownloader.log");
                }
            }

            try
            {
                if (File.Exists(_logFilePath))
                {
                    _currentLogSize = new FileInfo(_logFilePath).Length;
                }
            }
            catch { _currentLogSize = 0; }

            _fileWriterTask = Task.Run(FileWriterLoopAsync);
        }

        public void Info(string message, string category = "") => Log(LogLevel.INFO, message, category);
        public void Warning(string message, string category = "") => Log(LogLevel.WARN, message, category);
        public void Error(string message, string category = "") => Log(LogLevel.ERROR, message, category);
        public void Debug(string message, string category = "") => Log(LogLevel.DEBUG, message, category);
        public void Trace(string message, string category = "") => Log(LogLevel.TRACE, message, category);
        public void Fatal(string message, string category = "") => Log(LogLevel.FATAL, message, category);

        public void Log(LogLevel level, string message, string category = "")
        {
            if (level < _minLogLevel || level == LogLevel.OFF) return;
            var levelName = LevelNames.TryGetValue(level, out var n) ? n : "UNKNOWN";
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
                if (_disposed) return;
                try
                {
                    LogUpdated?.Invoke(this, formattedMessage);
                    System.Diagnostics.Debug.WriteLine(formattedMessage);
                    _pendingWrites.Enqueue(formattedMessage);
                    _writeSignal.Release();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ERROR] 日志记录失败: {ex.Message}");
                }
            }
        }

        private async Task FileWriterLoopAsync()
        {
            while (!_cts.Token.IsCancellationRequested)
            {
                try
                {
                    await Task.WhenAny(
                        _writeSignal.WaitAsync(_cts.Token),
                        Task.Delay(BufferFlushIntervalMs, _cts.Token)
                    );
                }
                catch (OperationCanceledException) { break; }
                FlushPendingWrites();
            }
            FlushPendingWrites();
        }

        private void FlushPendingWrites()
        {
            if (_pendingWrites.IsEmpty) return;
            lock (_fileLock)
            {
                if (_disposed) return;
                try
                {
                    if (_currentLogSize > MaxLogFileSize) RotateLogFile();
                    using var writer = new StreamWriter(_logFilePath, append: true, Encoding.UTF8);
                    while (_pendingWrites.TryDequeue(out var msg))
                    {
                        writer.WriteLine(msg);
                        _currentLogSize += Encoding.UTF8.GetByteCount(msg) + Environment.NewLine.Length;
                    }
                    writer.Flush();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ERROR] 日志文件写入失败: {ex.Message}");
                }
            }
        }

        private void RotateLogFile()
        {
            try
            {
                var oldestBackup = _logFilePath + "." + MaxBackupFiles;
                if (File.Exists(oldestBackup)) File.Delete(oldestBackup);
                for (int i = MaxBackupFiles - 1; i >= 1; i--)
                {
                    var src = _logFilePath + "." + i;
                    var dst = _logFilePath + "." + (i + 1);
                    if (File.Exists(src))
                    {
                        if (File.Exists(dst)) File.Delete(dst);
                        File.Move(src, dst);
                    }
                }
                if (File.Exists(_logFilePath)) File.Move(_logFilePath, _logFilePath + ".1");
                _currentLogSize = 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 日志滚动失败: {ex.Message}");
            }
        }

        public void SetLogLevel(LogLevel minLevel)
        {
            _minLogLevel = minLevel;
            Log(LogLevel.INFO, $"日志级别已设置为: {LevelNames[minLevel]}", "LoggingService");
        }

        public LogLevel GetCurrentLogLevel() => _minLogLevel;
        public string GetLogFilePath() => _logFilePath;

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            lock (_lockObject) { LogUpdated = null; }
            try
            {
                _cts.Cancel();
                if (_fileWriterTask != null && !_fileWriterTask.IsCompleted)
                    _fileWriterTask.Wait(2000);
                FlushPendingWrites();
            }
            catch { }
            finally
            {
                _cts.Dispose();
                _writeSignal.Dispose();
            }
        }
    }
}
