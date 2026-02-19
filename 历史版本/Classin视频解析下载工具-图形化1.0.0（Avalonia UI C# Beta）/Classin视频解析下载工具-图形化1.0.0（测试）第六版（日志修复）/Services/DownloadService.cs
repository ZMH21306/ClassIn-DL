using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Classin视频解析下载工具.Constants;

namespace Classin视频解析下载工具.Services
{
    /// <summary>
    /// 下载服务接口
    /// 提供文件下载、并发控制等功能
    /// </summary>
    public interface IDownloadService : IDisposable
    {
        /// <summary>
        /// 获取文件大小
        /// </summary>
        /// <param name="url">文件URL</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <returns>文件大小（字节）</returns>
        Task<long> GetFileSizeAsync(string url, CancellationToken cancellationToken = default);

        /// <summary>
        /// 发送HEAD请求获取文件信息
        /// </summary>
        /// <param name="url">文件URL</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <returns>HTTP响应消息</returns>
        Task<HttpResponseMessage> GetHeadAsync(string url, CancellationToken cancellationToken = default);

        /// <summary>
        /// 下载文件
        /// </summary>
        /// <param name="url">文件URL</param>
        /// <param name="filePath">保存路径</param>
        /// <param name="progressCallback">进度回调</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <returns>是否下载成功</returns>
        Task<bool> DownloadFileAsync(string url, string filePath, Action<long, long, double, TimeSpan> progressCallback, CancellationToken cancellationToken = default);

        /// <summary>
        /// 取消下载
        /// </summary>
        /// <param name="url">文件URL</param>
        void CancelDownload(string url);

        /// <summary>
        /// 带并发控制的下载
        /// </summary>
        /// <param name="url">文件URL</param>
        /// <param name="filePath">保存路径</param>
        /// <param name="progressCallback">进度回调</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <param name="onDownloadStarted">下载开始回调</param>
        /// <returns>是否下载成功</returns>
        Task<bool> DownloadWithConcurrencyControl(string url, string filePath, Action<long, long, double, TimeSpan> progressCallback, CancellationToken cancellationToken, Action? onDownloadStarted = null);

        /// <summary>
        /// 设置最大并发下载数
        /// </summary>
        /// <param name="max">最大并发数</param>
        void SetMaxConcurrentDownloads(int max);

        /// <summary>
        /// 获取最大并发下载数
        /// </summary>
        /// <returns>最大并发数</returns>
        int GetMaxConcurrentDownloads();



        /// <summary>
        /// 是否已释放
        /// </summary>
        bool IsDisposed { get; }
    }

    /// <summary>
    /// 下载服务实现
    /// 提供文件下载、并发控制、断点续传等功能
    /// </summary>
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

        /// <summary>
        /// 是否已释放
        /// </summary>
        public bool IsDisposed => _disposed;



        /// <summary>
        /// 默认构造函数
        /// </summary>
        public DownloadService() : this(new ValidationService())
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="validationService">验证服务</param>
        public DownloadService(IValidationService validationService)
        {
            _validationService = validationService ?? throw new ArgumentNullException(nameof(validationService));
            _activeDownloads = new ConcurrentDictionary<string, CancellationTokenSource>(StringComparer.OrdinalIgnoreCase);
            _disposeSemaphore = new SemaphoreSlim(1, 1);
            _staticLock = new object();
            _maxConcurrentDownloads = AppConstants.MaxConcurrentDownloads;

            InitializeHttpClient();
        }

        /// <summary>
        /// 初始化HttpClient
        /// </summary>
        private void InitializeHttpClient()
        {
            _httpClientHandler = new HttpClientHandler
            {
                UseProxy = false,
                Proxy = null,
                MaxConnectionsPerServer = AppConstants.HttpMaxConnectionsPerServer,
                UseDefaultCredentials = false,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
                CookieContainer = new CookieContainer(),
                AllowAutoRedirect = true,
                MaxAutomaticRedirections = 5,
                UseCookies = true
            };

            _httpClient = new HttpClient(_httpClientHandler)
            {
                Timeout = TimeSpan.FromHours(AppConstants.DefaultTimeoutHours)
            };
            
            _httpClient.DefaultRequestHeaders.ConnectionClose = false;
            _httpClient.DefaultRequestHeaders.CacheControl = new CacheControlHeaderValue { NoCache = true, MaxAge = TimeSpan.FromSeconds(0) };
            _httpClient.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate, br");
            _httpClient.DefaultRequestHeaders.Add("Accept-Language", "zh-CN,zh;q=0.9,en;q=0.8");
            _httpClient.DefaultRequestHeaders.Add("Connection", "keep-alive");

            ServicePointManager.DefaultConnectionLimit = AppConstants.HttpMaxConnectionsPerServer;
            ServicePointManager.ReusePort = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
            ServicePointManager.MaxServicePointIdleTime = AppConstants.ServicePointIdleTimeMs;
            ServicePointManager.Expect100Continue = false;
            ServicePointManager.UseNagleAlgorithm = false;

            _concurrencySemaphore = new SemaphoreSlim(_maxConcurrentDownloads, _maxConcurrentDownloads);
        }

        /// <summary>
        /// 设置最大并发下载数
        /// </summary>
        /// <param name="max">最大并发数</param>
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

        /// <summary>
        /// 获取最大并发下载数
        /// </summary>
        /// <returns>最大并发数</returns>
        public int GetMaxConcurrentDownloads() => _maxConcurrentDownloads;

        /// <summary>
        /// 获取文件大小
        /// </summary>
        /// <param name="url">文件URL</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <returns>文件大小（字节）</returns>
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

        /// <summary>
        /// 发送HEAD请求获取文件信息
        /// </summary>
        /// <param name="url">文件URL</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <returns>HTTP响应消息</returns>
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

        /// <summary>
        /// 配置请求头
        /// </summary>
        /// <param name="request">HTTP请求消息</param>
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

        /// <summary>
        /// 下载文件
        /// </summary>
        /// <param name="url">文件URL</param>
        /// <param name="filePath">保存路径</param>
        /// <param name="progressCallback">进度回调</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <returns>是否下载成功</returns>
        public async Task<bool> DownloadFileAsync(string url, string filePath, Action<long, long, double, TimeSpan> progressCallback, CancellationToken cancellationToken = default)
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

                    // 自动计算最优线程数（基于CPU核心数）
                    var optimalThreadCount = Math.Max(Environment.ProcessorCount, 1);
                    var bufferSize = CalculateOptimalBufferSize(totalBytesValue);
                    var bufferOwner = System.Buffers.ArrayPool<byte>.Shared.Rent(bufferSize);
                    var buffer = new Memory<byte>(bufferOwner, 0, bufferSize);
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

                            var bytesRead = await contentStream.ReadAsync(buffer, cancellationToken);
                            if (bytesRead == 0) break;

                            await fileStream.WriteAsync(buffer.Slice(0, bytesRead), cancellationToken);

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

                    System.Buffers.ArrayPool<byte>.Shared.Return(bufferOwner);

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

        /// <summary>
        /// 计算最优缓冲区大小
        /// </summary>
        /// <param name="totalBytes">文件总大小</param>
        /// <returns>缓冲区大小</returns>
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

        /// <summary>
        /// 计算指数退避延迟
        /// </summary>
        /// <param name="retryCount">重试次数</param>
        /// <returns>延迟时间（毫秒）</returns>
        private static int CalculateExponentialBackoff(int retryCount)
        {
            var delay = (int)(AppConstants.ExponentialBackoffBaseMs * Math.Pow(2, retryCount - 1));
            return Math.Min(delay, AppConstants.ExponentialBackoffMaxMs);
        }

        /// <summary>
        /// 取消下载
        /// </summary>
        /// <param name="url">文件URL</param>
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

        /// <summary>
        /// 带并发控制的下载
        /// </summary>
        /// <param name="url">文件URL</param>
        /// <param name="filePath">保存路径</param>
        /// <param name="progressCallback">进度回调</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <param name="onDownloadStarted">下载开始回调</param>
        /// <returns>是否下载成功</returns>
        public async Task<bool> DownloadWithConcurrencyControl(string url, string filePath, Action<long, long, double, TimeSpan> progressCallback, CancellationToken cancellationToken, Action? onDownloadStarted = null)
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

                    return await DownloadFileAsync(url, filePath, progressCallback, cts.Token);
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

        /// <summary>
        /// 释放资源
        /// </summary>
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