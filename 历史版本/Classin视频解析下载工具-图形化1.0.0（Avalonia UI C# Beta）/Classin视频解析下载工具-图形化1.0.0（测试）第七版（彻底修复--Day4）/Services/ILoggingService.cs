using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Classin视频解析下载工具.Services
{
    /// <summary>
    /// 日志级别枚举
    /// </summary>
    public enum LogLevel
    {
        ALL = 0,
        TRACE = 1,
        DEBUG = 2,
        INFO = 3,
        WARN = 4,
        ERROR = 5,
        FATAL = 6,
        OFF = 7
    }

    public interface ILoggingService : IDisposable
    {
        // 现有方法保持兼容
        void Info(string message, string category = "");
        void Warning(string message, string category = "");
        void Error(string message, string category = "");
        void Debug(string message, string category = "");
        
        // 新增方法
        void Trace(string message, string category = "");
        void Fatal(string message, string category = "");
        void Log(LogLevel level, string message, string category = "");
        
        // 配置方法
        void SetLogLevel(LogLevel minLevel);
        LogLevel GetCurrentLogLevel();
        
        event EventHandler<string>? LogUpdated;
    }
}