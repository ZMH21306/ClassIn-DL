using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Classin视频解析下载工具.Services
{
    public class MemoryMonitor : IMemoryMonitor
    {
        private readonly Process _currentProcess;
        private readonly List<MemorySnapshot> _snapshots;
        private readonly object _snapshotLock;
        private CancellationTokenSource _cts;
        private Task _monitoringTask;
        private bool _isMonitoring;

        public MemoryMonitor()
        {
            _currentProcess = Process.GetCurrentProcess();
            _snapshots = new List<MemorySnapshot>();
            _snapshotLock = new object();
            _cts = new CancellationTokenSource();
        }

        public void StartMonitoring()
        {
            if (_isMonitoring) return;

            _cts = new CancellationTokenSource();
            _isMonitoring = true;

            _monitoringTask = Task.Run(async () =>
            {
                while (!_cts.Token.IsCancellationRequested)
                {
                    TakeSnapshot($"Auto-{DateTime.Now:HH:mm:ss}");
                    await Task.Delay(30000, _cts.Token); // 每30秒自动拍摄一次快照
                }
            }, _cts.Token);
        }

        public void StopMonitoring()
        {
            if (!_isMonitoring) return;

            _cts.Cancel();
            _monitoringTask?.Wait(1000);
            _isMonitoring = false;
        }

        public MemoryUsage GetCurrentMemoryUsage()
        {
            _currentProcess.Refresh();

            return new MemoryUsage
            {
                PrivateMemorySizeMb = _currentProcess.PrivateMemorySize64 / 1024.0 / 1024.0,
                WorkingSetMb = _currentProcess.WorkingSet64 / 1024.0 / 1024.0,
                PagedMemorySizeMb = _currentProcess.PagedMemorySize64 / 1024.0 / 1024.0,
                Timestamp = DateTime.Now
            };
        }

        public List<MemorySnapshot> GetMemorySnapshots()
        {
            lock (_snapshotLock)
            {
                return new List<MemorySnapshot>(_snapshots);
            }
        }

        public void TakeSnapshot(string label)
        {
            var usage = GetCurrentMemoryUsage();
            var snapshot = new MemorySnapshot
            {
                Label = label,
                Usage = usage,
                Timestamp = DateTime.Now
            };

            lock (_snapshotLock)
            {
                _snapshots.Add(snapshot);
                
                // 限制快照数量，最多保存50个
                if (_snapshots.Count > 50)
                {
                    _snapshots.RemoveAt(0);
                }
            }
        }

        public double CalculateMemoryGrowth()
        {
            lock (_snapshotLock)
            {
                if (_snapshots.Count < 2)
                    return 0;

                var first = _snapshots[0].Usage.PrivateMemorySizeMb;
                var last = _snapshots[_snapshots.Count - 1].Usage.PrivateMemorySizeMb;
                
                return last - first;
            }
        }

        public bool IsMemoryLeakDetected()
        {
            lock (_snapshotLock)
            {
                if (_snapshots.Count < 3)
                    return false;

                // 检查最近3次快照的内存使用趋势
                var recentSnapshots = _snapshots.GetRange(_snapshots.Count - 3, 3);
                
                // 如果内存持续增长，可能存在泄漏
                for (int i = 1; i < recentSnapshots.Count; i++)
                {
                    if (recentSnapshots[i].Usage.PrivateMemorySizeMb < recentSnapshots[i - 1].Usage.PrivateMemorySizeMb)
                    {
                        return false;
                    }
                }

                // 检查内存增长幅度
                var growth = CalculateMemoryGrowth();
                return growth > 50; // 如果增长超过50MB，认为可能存在泄漏
            }
        }

        public void Dispose()
        {
            StopMonitoring();
            _cts.Dispose();
        }
    }
}
