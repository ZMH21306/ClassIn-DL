using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Classin视频解析下载工具.Services
{
    public class MemoryUsage
    {
        public double PrivateMemorySizeMb { get; set; }
        public double WorkingSetMb { get; set; }
        public double PagedMemorySizeMb { get; set; }
        public DateTime Timestamp { get; set; }

        public override string ToString()
        {
            return $"Private: {PrivateMemorySizeMb:F2} MB, Working Set: {WorkingSetMb:F2} MB, Timestamp: {Timestamp:HH:mm:ss}";
        }
    }

    public class MemorySnapshot
    {
        public string Label { get; set; }
        public MemoryUsage Usage { get; set; }
        public DateTime Timestamp { get; set; }

        public override string ToString()
        {
            return $"[{Label}] {Usage}";
        }
    }

    public interface IMemoryMonitor
    {
        void StartMonitoring();
        void StopMonitoring();
        MemoryUsage GetCurrentMemoryUsage();
        List<MemorySnapshot> GetMemorySnapshots();
        void TakeSnapshot(string label);
        double CalculateMemoryGrowth();
        bool IsMemoryLeakDetected();
    }
}