namespace VideoDownloader.Constants
{
    public static class AppConstants
    {
        public const int DefaultMaxRetries = 3;
        public const int DownloadServiceMaxRetries = 5;
        public const int DefaultBufferSize = 1024 * 1024;
        public const int MinBufferSize = 64 * 1024;
        public const int MaxBufferSize = 4 * 1024 * 1024;
        public const int UiUpdateThrottleMs = 100;
        public const int DefaultTimeoutHours = 6;
        public const int DefaultTimeoutSeconds = 21600;
        public const int MaxConcurrentDownloads = 5;
        public const int MaxDownloadThreads = 32;
        public const int ThreadClampMin = 1;
        public const int ThreadClampMax = 256;
        public const int ExponentialBackoffMaxMs = 60000;
        public const int ExponentialBackoffBaseMs = 1000;
        public const int ProgressUpdateIntervalMs = 200;
        public const int FileStreamBufferSize = 1024 * 1024;
        public const int HttpMaxConnectionsPerServer = 100;
        public const int ServicePointIdleTimeMs = 10000;
        public const int ClipboardRetryCount = 5;
        public const int ClipboardRetryDelayMs = 100;
        public const string DefaultUserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
        public const string DefaultReferrer = "https://www.eeo.cn/";
        public const string AcceptedContentTypes = "video/mp4, */*";
        public const string DefaultDownloadFolder = "下载目录";
        public const string LogFileName = "VideoDownloader.log";
        public const string ConfigFileName = "appsettings.json";
        public static readonly string[] AllowedVideoExtensions = { ".mp4", ".mkv", ".avi", ".mov", ".flv", ".wmv" };
        public static readonly string[] AllowedDomains = { "eeo.cn", "classin.com", "classin.tech" };
    }
}