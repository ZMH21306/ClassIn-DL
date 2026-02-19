using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using Classin视频解析下载工具.Services;

namespace Classin视频解析下载工具.ViewModels
{
    public class LogViewModel : ViewModelBase, IDisposable
    {
        private readonly StringBuilder _logBuffer;
        private readonly object _lockObject = new();
        private bool _isInitialized;
        private bool _disposed;
        private int _throttleCounter = 0;
        private const int THROTTLE_THRESHOLD = 10; // 每10条日志更新一次UI
        private const int MAX_LOG_LINES = 2000; // 最大日志行数
        private const int TRIM_TO_LINES = 1000; // 截断到的行数
        private Timer? _uiUpdateTimer;
        private bool _pendingUiUpdate = false;
        
        // 日志过滤
        private LogLevel _minLogLevel = LogLevel.INFO;
        private string _filterText = string.Empty;
        private List<string> _filteredLogLines = new();

        public LogViewModel(ILoggingService? loggingService = null)
        {
            _logBuffer = new StringBuilder();
            // 创建定时器，定期更新UI以避免频繁刷新
            _uiUpdateTimer = new Timer(OnUiUpdateTimer, null, Timeout.Infinite, Timeout.Infinite);
            
            // 如果提供了LoggingService，则订阅其事件
            if (loggingService != null)
            {
                loggingService.LogUpdated += OnLogUpdated;
            }
        }

        public LogLevel MinLogLevel
        {
            get => _minLogLevel;
            set
            {
                if (SetProperty(ref _minLogLevel, value))
                {
                    RefreshFilteredLogs();
                }
            }
        }

        public string FilterText
        {
            get => _filterText;
            set
            {
                if (SetProperty(ref _filterText, value))
                {
                    RefreshFilteredLogs();
                }
            }
        }

        private string _filteredLogText = string.Empty;
        public string FilteredLogText
        {
            get => _filteredLogText;
            private set => SetProperty(ref _filteredLogText, value);
        }

        public bool IsInitialized
        {
            get => _isInitialized;
            set => SetProperty(ref _isInitialized, value);
        }

        private void OnLogUpdated(object? sender, string logMessage)
        {
            AppendLogDirect(logMessage);
        }

        private void AppendLogDirect(string message)
        {
            // 允许在初始化前记录关键日志，但不在UI中显示
            if (_disposed) 
            {
                return;
            }
            
            // 如果未初始化，仍然记录到缓冲区但不更新UI
            bool shouldUpdateUI = _isInitialized;

            lock (_lockObject)
            {
                _logBuffer.AppendLine(message);
                _filteredLogLines.Add(message);
                
                ApplyFilters();
                
                if (_filteredLogLines.Count > MAX_LOG_LINES)
                {
                    TrimLogs();
                }

                // 只有在初始化后才更新UI
                if (shouldUpdateUI)
                {
                    ScheduleUiUpdate();
                }
            }
        }

        public void AppendLog(string message)
        {
            if (!_isInitialized || _disposed) return;

            lock (_lockObject)
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var logEntry = $"[{timestamp}] {message}";
                
                _logBuffer.AppendLine(logEntry);
                _filteredLogLines.Add(logEntry);
                
                // 应用过滤
                ApplyFilters();
                
                // 限制日志缓冲区大小
                if (_filteredLogLines.Count > MAX_LOG_LINES)
                {
                    TrimLogs();
                }

                // 节流更新UI
                _throttleCounter++;
                if (_throttleCounter >= 1) // 立即更新用于测试
                {
                    ScheduleUiUpdate();
                    _throttleCounter = 0;
                }
            }
        }

        private void ApplyFilters()
        {
            // 目前简单的过滤实现
            // 可以在这里添加更复杂的过滤逻辑
        }

        private void TrimLogs()
        {
            // 保留最新的日志行
            var startIndex = Math.Max(0, _filteredLogLines.Count - TRIM_TO_LINES);
            _filteredLogLines = _filteredLogLines.GetRange(startIndex, _filteredLogLines.Count - startIndex);
            
            // 更新缓冲区
            _logBuffer.Clear();
            foreach (var line in _filteredLogLines)
            {
                _logBuffer.AppendLine(line);
            }
            _logBuffer.AppendLine("--- 日志已自动截断 ---");
        }

        private void ScheduleUiUpdate()
        {
            if (!_pendingUiUpdate)
            {
                _pendingUiUpdate = true;
                _uiUpdateTimer?.Change(1, Timeout.Infinite); // 1ms后立即更新
            }
        }

        private void OnUiUpdateTimer(object? state)
        {
            lock (_lockObject)
            {
                if (_pendingUiUpdate && !_disposed)
                {
                    var newText = string.Join(Environment.NewLine, _filteredLogLines);
                    FilteredLogText = newText;
                    _pendingUiUpdate = false;
                }
            }
        }

        private void RefreshFilteredLogs()
        {
            lock (_lockObject)
            {
                ApplyFilters();
                ScheduleUiUpdate();
            }
        }



        public void ClearLog()
        {
            lock (_lockObject)
            {
                _logBuffer.Clear();
                _filteredLogLines.Clear();
                FilteredLogText = string.Empty;
                _throttleCounter = 0;
                _pendingUiUpdate = false;
            }
        }

        public string GetLogContent()
        {
            lock (_lockObject)
            {
                return _logBuffer.ToString();
            }
        }

        public string GetFilteredLogContent()
        {
            lock (_lockObject)
            {
                return string.Join(Environment.NewLine, _filteredLogLines);
            }
        }

        public int GetLogLineCount()
        {
            lock (_lockObject)
            {
                return _filteredLogLines.Count;
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
            catch (Exception ex)
            {
                // 记录保存错误但不抛出异常
                System.Diagnostics.Debug.WriteLine($"保存日志文件失败: {ex.Message}");
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            
            // 停止定时器
            _uiUpdateTimer?.Change(Timeout.Infinite, Timeout.Infinite);
            _uiUpdateTimer?.Dispose();
            _uiUpdateTimer = null;
            
            // 立即更新UI以显示最终状态
            OnUiUpdateTimer(null);
            
            // 保存日志到文件
            try
            {
                var logDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                                        "Classin视频解析下载工具", "Logs");
                System.IO.Directory.CreateDirectory(logDir);
                
                var logFile = System.IO.Path.Combine(logDir, $"app_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                SaveLogToFile(logFile);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"保存日志文件时发生异常: {ex.Message}");
            }
            
            lock (_lockObject)
            {
                _filteredLogLines.Clear();
                _logBuffer.Clear();
            }
        }
    }
}