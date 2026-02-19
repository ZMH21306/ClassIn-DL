using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using VideoDownloader.Constants;

namespace VideoDownloader.Services
{
    public interface IDownloadService : IDisposable
    {
        Task<long> GetFileSizeAsync(string url, CancellationToken cancellationToken = default);
        Task<HttpResponseMessage> GetHeadAsync(string url, CancellationToken cancellationToken = default);
        Task<bool> DownloadFileAsync(string url, string filePath, Action<long, long, double, TimeSpan> progressCallback, int threads, CancellationToken cancellationToken = default);
        void CancelDownload(string url);
        Task<bool> DownloadWithConcurrencyControl(string url, string filePath, Action<long, long, double, TimeSpan> progressCallback, CancellationToken cancellationToken, int threads, Action? onDownloadStarted = null);
        void SetMaxConcurrentDownloads(int max);
        int GetMaxConcurrentDownloads();
        int MaxDownloadThreads { get; set; }
        bool IsDisposed { get; }
    }

    public sealed class DownloadService : IDownloadService
    {
        private HttpClient _httpClient = null!;
        private HttpClientHandler _httpClientHandler = null!;
        private readonly ConcurrentDictionary<string, CancellationTokenSource> _activeDownloads;
        private SemaphoreSlim _concurrencySemaphore = null!;
        private readonly SemaphoreSlim _disposeSemaphore;
        private readonly IValidationService _validationService;
        private readonly object _staticLock;
        private volatile bool _disposed;
        private volatile int _maxConcurrentDownloads;
        private volatile int _maxDownloadThreads;

        public bool IsDisposed => _disposed;

        public int MaxDownloadThreads
        {
            get => _maxDownloadThreads;
            set => _maxDownloadThreads = Math.Clamp(value, AppConstants.ThreadClampMin, AppConstants.ThreadClampMax);
        }

        public DownloadService() : this(new ValidationService())
        {
        }

        public DownloadService(IValidationService validationService)
        {
            _validationService = validationService ?? throw new ArgumentNullException(nameof(validationService));
            _activeDownloads = new ConcurrentDictionary<string, CancellationTokenSource>(StringComparer.OrdinalIgnoreCase);
            _disposeSemaphore = new SemaphoreSlim(1, 1);
            _staticLock = new object();
            _maxConcurrentDownloads = AppConstants.MaxConcurrentDownloads;
            _maxDownloadThreads = AppConstants.MaxDownloadThreads;

            InitializeHttpClient();
        }

        private void InitializeHttpClient()
        {
            _httpClientHandler = new HttpClientHandler
            {
                UseProxy = false,
                Proxy = null,
                MaxConnectionsPerServer = AppConstants.HttpMaxConnectionsPerServer,
                UseDefaultCredentials = false,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
                CookieContainer = new CookieContainer()
            };

            _httpClient = new HttpClient(_httpClientHandler)
            {
                Timeout = TimeSpan.FromHours(AppConstants.DefaultTimeoutHours),
                DefaultRequestHeaders =
                {
                    ConnectionClose = false,
                    CacheControl = new CacheControlHeaderValue { NoCache = true }
                }
            };

            ServicePointManager.DefaultConnectionLimit = AppConstants.HttpMaxConnectionsPerServer;
            ServicePointManager.ReusePort = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
            ServicePointManager.MaxServicePointIdleTime = AppConstants.ServicePointIdleTimeMs;

            _concurrencySemaphore = new SemaphoreSlim(_maxConcurrentDownloads, _maxConcurrentDownloads);
        }

        public void SetMaxConcurrentDownloads(int max)
        {
            if (max <= 0) max = 1;
            if (max > 100) max = 100;

            lock (_staticLock)
            {
                if (_disposed) return;

                var oldMax = _maxConcurrentDownloads;
                _maxConcurrentDownloads = max;

                if (max > oldMax)
                {
                    var releaseCount = max - oldMax;
                    try
                    {
                        _concurrencySemaphore.Release(releaseCount);
                    }
                    catch (SemaphoreFullException)
                    {
                        Debug.WriteLine("信号量已达到最大计数，无法释放");
                    }
                    catch (ObjectDisposedException)
                    {
                        Debug.WriteLine("信号量已释放，无法操作");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"释放信号量失败: {ex.Message}");
                    }
                }
                else if (max < oldMax)
                {
                    var diff = oldMax - max;
                    for (int i = 0; i < diff; i++)
                    {
                        if (_concurrencySemaphore.CurrentCount > 0)
                        {
                            try
                            {
                                _concurrencySemaphore.Wait(0);
                            }
                            catch (OperationCanceledException)
                            {
                                break;
                            }
                            catch (ObjectDisposedException)
                            {
                                break;
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"等待信号量失败: {ex.Message}");
                                break;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
        }

        public int GetMaxConcurrentDownloads() => _maxConcurrentDownloads;

        public async Task<long> GetFileSizeAsync(string url, CancellationToken cancellationToken = default)
        {
            var validationResult = _validationService.ValidateUrl(url);
            if (!validationResult.IsValid)
            {
                Debug.WriteLine($"URL验证失败: {validationResult.ErrorMessage}");
                return 0;
            }

            try
            {
                using var response = await GetHeadAsync(url, cancellationToken);
                if (response.IsSuccessStatusCode)
                {
                    return response.Content.Headers.ContentLength ?? 0;
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取文件大小失败: {ex.Message}");
            }

            return 0;
        }

        public async Task<HttpResponseMessage> GetHeadAsync(string url, CancellationToken cancellationToken = default)
        {
            var validationResult = _validationService.ValidateUrl(url);
            if (!validationResult.IsValid)
            {
                return new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent($"URL验证失败: {validationResult.ErrorMessage}")
                };
            }

            try
            {
                using var request = new HttpRequestMessage(HttpMethod.Head, url);
                ConfigureRequestHeaders(request);

                return await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"HEAD请求失败: {ex.Message}");
                return new HttpResponseMessage(HttpStatusCode.InternalServerError)
                {
                    Content = new StringContent($"请求失败: {ex.Message}")
                };
            }
        }

        private void ConfigureRequestHeaders(HttpRequestMessage request)
        {
            request.Headers.UserAgent.ParseAdd(AppConstants.DefaultUserAgent);
            request.Headers.Accept.ParseAdd(AppConstants.AcceptedContentTypes);

            if (!string.IsNullOrEmpty(AppConstants.DefaultReferrer))
            {
                try
                {
                    request.Headers.Referrer = new Uri(AppConstants.DefaultReferrer);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"设置 Referrer 头失败: {ex.Message}");
                }
            }
        }

        public async Task<bool> DownloadFileAsync(string url, string filePath, Action<long, long, double, TimeSpan> progressCallback, int threads, CancellationToken cancellationToken = default)
        {
            var urlValidation = _validationService.ValidateUrl(url);
            if (!urlValidation.IsValid)
            {
                Debug.WriteLine($"URL验证失败: {urlValidation.ErrorMessage}");
                return false;
            }

            var filePathValidation = _validationService.ValidateFilePath(filePath);
            if (!filePathValidation.IsValid)
            {
                Debug.WriteLine($"文件路径验证失败: {filePathValidation.ErrorMessage}");
                return false;
            }

            const int maxRetries = AppConstants.DownloadServiceMaxRetries;
            int retryCount = 0;
            long totalBytesRead = 0;
            long? totalBytes = null;
            long startPosition = 0;
            long totalBytesValue = 0;
            bool isFinalUpdateSent = false;
            Exception? lastException = null;

            while (retryCount <= maxRetries)
            {
                try
                {
                    if (totalBytes == null)
                    {
                        using var headResponse = await GetHeadAsync(url, cancellationToken);
                        if (headResponse.IsSuccessStatusCode)
                        {
                            totalBytes = headResponse.Content.Headers.ContentLength;
                        }
                        else if (headResponse.StatusCode == HttpStatusCode.NotFound)
                        {
                            throw new FileNotFoundException($"文件不存在: {url}");
                        }
                    }

                    totalBytesValue = totalBytes ?? 0;

                    var directory = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    using var request = new HttpRequestMessage(HttpMethod.Get, url);
                    ConfigureRequestHeaders(request);

                    if (startPosition > 0)
                    {
                        request.Headers.Range = new RangeHeaderValue(startPosition, null);
                    }

                    using var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);

                    if (!response.IsSuccessStatusCode && response.StatusCode != HttpStatusCode.PartialContent)
                    {
                        throw new HttpRequestException($"服务器返回错误状态: {(int)response.StatusCode} {response.StatusCode}");
                    }

                    if (response.StatusCode == HttpStatusCode.PartialContent)
                    {
                        var contentLength = response.Content.Headers.ContentLength;
                        if (contentLength.HasValue)
                        {
                            totalBytes = startPosition + contentLength.Value;
                            totalBytesValue = totalBytes.Value;
                        }
                    }
                    else if (!totalBytes.HasValue)
                    {
                        totalBytes = response.Content.Headers.ContentLength;
                        totalBytesValue = totalBytes ?? 0;
                    }

                    using var contentStream = await response.Content.ReadAsStreamAsync(cancellationToken);

                    var bufferSize = CalculateOptimalBufferSize(totalBytesValue);
                    var buffer = new byte[bufferSize];
                    var lastBytesRead = startPosition;
                    var lastUpdateTime = Stopwatch.StartNew();
                    var progressThrottle = TimeSpan.FromMilliseconds(AppConstants.ProgressUpdateIntervalMs);

                    using (var fileStream = new FileStream(
                        filePath,
                        startPosition > 0 ? FileMode.Append : FileMode.Create,
                        FileAccess.Write,
                        FileShare.None,
                        AppConstants.FileStreamBufferSize,
                        FileOptions.Asynchronous | FileOptions.SequentialScan))
                    {
                        while (true)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            var bytesRead = await contentStream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken);
                            if (bytesRead == 0) break;

                            await fileStream.WriteAsync(buffer.AsMemory(0, bytesRead), cancellationToken);

                            totalBytesRead += bytesRead;

                            if (lastUpdateTime.Elapsed > progressThrottle)
                            {
                                var speed = (totalBytesRead - lastBytesRead) / lastUpdateTime.Elapsed.TotalSeconds;
                                lastBytesRead = totalBytesRead;
                                lastUpdateTime.Restart();

                                var remainingTime = totalBytesValue > 0
                                    ? TimeSpan.FromSeconds((totalBytesValue - totalBytesRead) / (speed > 0 ? speed : 1))
                                    : TimeSpan.MaxValue;

                                progressCallback?.Invoke(totalBytesRead, totalBytesValue, speed, remainingTime);
                            }
                        }

                        await fileStream.FlushAsync(cancellationToken);
                    }

                    progressCallback?.Invoke(totalBytesValue, totalBytesValue, 0, TimeSpan.Zero);
                    isFinalUpdateSent = true;

                    await Task.Delay(AppConstants.ProgressUpdateIntervalMs, cancellationToken);

                    return true;
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (HttpRequestException ex) when (ex.StatusCode.HasValue && (int)ex.StatusCode.Value >= 500 && retryCount < maxRetries)
                {
                    lastException = ex;
                    retryCount++;
                    var delay = CalculateExponentialBackoff(retryCount);
                    await Task.Delay(delay, cancellationToken);
                }
                catch (IOException ioEx) when (ioEx is FileNotFoundException || ioEx is DirectoryNotFoundException || ioEx.HResult == -2147024894)
                {
                    throw;
                }
                catch (Exception ex) when (retryCount < maxRetries)
                {
                    lastException = ex;
                    Debug.WriteLine($"下载失败，正在重试({retryCount}/{maxRetries}): {ex.Message}");
                    retryCount++;
                    var delay = CalculateExponentialBackoff(retryCount);
                    await Task.Delay(delay, cancellationToken);
                }
                finally
                {
                    if (!isFinalUpdateSent)
                    {
                        try
                        {
                            progressCallback?.Invoke(totalBytesRead, totalBytesValue, 0, TimeSpan.Zero);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"进度回调失败: {ex.Message}");
                        }
                        isFinalUpdateSent = true;
                    }
                }
            }

            if (lastException != null)
            {
                throw new AggregateException($"下载失败，已重试{maxRetries}次", lastException);
            }

            return false;
        }

        private int CalculateOptimalBufferSize(long totalBytes)
        {
            if (totalBytes >= 100 * 1024 * 1024)
            {
                return AppConstants.MaxBufferSize;
            }

            if (totalBytes >= 10 * 1024 * 1024)
            {
                return 2 * 1024 * 1024;
            }

            if (totalBytes >= 1024 * 1024)
            {
                return AppConstants.DefaultBufferSize;
            }

            return 256 * 1024;
        }

        private static int CalculateExponentialBackoff(int retryCount)
        {
            var delay = (int)(AppConstants.ExponentialBackoffBaseMs * Math.Pow(2, retryCount - 1));
            return Math.Min(delay, AppConstants.ExponentialBackoffMaxMs);
        }

        public void CancelDownload(string url)
        {
            if (_disposed) return;

            lock (_staticLock)
            {
                if (_disposed) return;

                if (_activeDownloads.TryRemove(url, out var cts))
                {
                    try
                    {
                        cts.Cancel();
                        cts.Dispose();
                    }
                    catch
                    {
                    }
                }
            }
        }

        public async Task<bool> DownloadWithConcurrencyControl(string url, string filePath, Action<long, long, double, TimeSpan> progressCallback, CancellationToken cancellationToken, int threads, Action? onDownloadStarted = null)
        {
            if (_disposed)
            {
                return false;
            }

            await _disposeSemaphore.WaitAsync(cancellationToken);
            try
            {
                if (_disposed)
                {
                    return false;
                }

                var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

                lock (_staticLock)
                {
                    if (_disposed)
                    {
                        cts.Dispose();
                        return false;
                    }
                    _activeDownloads[url] = cts;
                }

                try
                {
                    await _concurrencySemaphore.WaitAsync(cts.Token);
                    cancellationToken.ThrowIfCancellationRequested();

                    onDownloadStarted?.Invoke();

                    return await DownloadFileAsync(url, filePath, progressCallback, threads, cts.Token);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                finally
                {
                    lock (_staticLock)
                    {
                        _activeDownloads.TryRemove(url, out _);
                    }

                    try
                    {
                        _concurrencySemaphore.Release();
                    }
                    catch
                    {
                    }

                    cts.Dispose();
                }
            }
            finally
            {
                _disposeSemaphore.Release();
            }
        }

        public void Dispose()
        {
            if (_disposed) return;

            _disposeSemaphore.Wait();
            try
            {
                if (_disposed) return;
                _disposed = true;

                lock (_staticLock)
                {
                    foreach (var kvp in _activeDownloads)
                    {
                        try
                        {
                            kvp.Value.Cancel();
                            kvp.Value.Dispose();
                        }
                        catch
                        {
                        }
                    }
                    _activeDownloads.Clear();
                }

                _concurrencySemaphore?.Dispose();
                _httpClient?.Dispose();
                _httpClientHandler?.Dispose();
            }
            finally
            {
                _disposeSemaphore.Release();
            }
        }
    }
}