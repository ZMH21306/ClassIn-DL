using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Classin视频解析下载工具.Services.SystemServices
{
    public interface INetworkMonitor : IDisposable
    {
        bool IsNetworkAvailable { get; }
        event EventHandler<bool>? NetworkAvailabilityChanged;
        void Stop();
    }

    public class NetworkMonitor : INetworkMonitor
    {
        private bool _isNetworkAvailable;
        private readonly System.Threading.Timer _timer;
        private bool _disposed;

        public bool IsNetworkAvailable => _isNetworkAvailable;

        public event EventHandler<bool>? NetworkAvailabilityChanged;

        public NetworkMonitor()
        {
            _isNetworkAvailable = CheckNetworkConnectivity();
            _timer = new System.Threading.Timer(CheckNetworkStatus, null, TimeSpan.Zero, TimeSpan.FromSeconds(30));
        }

        private void CheckNetworkStatus(object? state)
        {
            var currentStatus = CheckNetworkConnectivity();
            if (currentStatus != _isNetworkAvailable)
            {
                _isNetworkAvailable = currentStatus;
                NetworkAvailabilityChanged?.Invoke(this, _isNetworkAvailable);
            }
        }

        private static bool CheckNetworkConnectivity()
        {
            try
            {
                // 检查网络连接
                return System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();
            }
            catch
            {
                return false;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                _timer?.Dispose();
                NetworkAvailabilityChanged = null;
            }
            catch (Exception ex)
            {
                // 记录异常但不影响退出
                Console.WriteLine($"NetworkMonitor清理异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 停止网络监控
        /// </summary>
        public void Stop()
        {
            Dispose();
        }
    }
}