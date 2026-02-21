using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using Classin视频解析下载工具.Shared.Commands;
using Classin视频解析下载工具.Models;
using Classin视频解析下载工具.Services.CoreServices;
using Classin视频解析下载工具.Services.DownloadServices;

namespace Classin视频解析下载工具.ViewModels
{
    public class VideoListViewModel : ViewModelBase
    {
        private readonly IDuplicateDetectionService _duplicateDetectionService;
        private ObservableCollection<VideoItem> _videoItems;
        private VideoItem? _selectedItem;
        private string _statusSummary = "当前没有视频正在下载，统计信息为空";

        public VideoListViewModel(IDuplicateDetectionService duplicateDetectionService)
        {
            _duplicateDetectionService = duplicateDetectionService;
            _videoItems = new ObservableCollection<VideoItem>();
            InitializeCommands();
        }

        public ObservableCollection<VideoItem> VideoItems
        {
            get => _videoItems;
            set => SetProperty(ref _videoItems, value);
        }

        public VideoItem? SelectedItem
        {
            get => _selectedItem;
            set => SetProperty(ref _selectedItem, value);
        }

        public string StatusSummary
        {
            get => _statusSummary;
            set => SetProperty(ref _statusSummary, value);
        }

        // Commands
        public RelayCommand<VideoItem> CopyUrlCommand { get; private set; } = null!;
        public RelayCommand<VideoItem> DownloadSingleCommand { get; private set; } = null!;
        public RelayCommand<VideoItem> DeleteItemCommand { get; private set; } = null!;

        private void InitializeCommands()
        {
            CopyUrlCommand = new RelayCommand<VideoItem>(ExecuteCopyUrlCommand);
            DownloadSingleCommand = new RelayCommand<VideoItem>(ExecuteDownloadSingleCommand);
            DeleteItemCommand = new RelayCommand<VideoItem>(ExecuteDeleteItemCommand);
        }

        private void ExecuteCopyUrlCommand(VideoItem? item)
        {
            // 这个方法会被MainViewModel中的实际实现覆盖
        }

        private void ExecuteDownloadSingleCommand(VideoItem? item)
        {
            // 这个方法会被MainViewModel中的实际实现覆盖
        }

        private void ExecuteDeleteItemCommand(VideoItem? item)
        {
            // 这个方法会被MainViewModel中的实际实现覆盖
        }

        public void AddVideoItem(VideoItem item)
        {
            if (!_duplicateDetectionService.IsDuplicate(item.Name))
            {
                _duplicateDetectionService.AddToCache(item);
                VideoItems.Add(item);
                UpdateVideoIndexes();
                UpdateStatusSummary();
            }
        }

        public void AddVideoItems(IEnumerable<VideoItem> items)
        {
            foreach (var item in items)
            {
                AddVideoItem(item);
            }
        }

        public void RemoveVideoItem(VideoItem item)
        {
            VideoItems.Remove(item);
            _duplicateDetectionService.RemoveFromCache(item.Name);
            UpdateVideoIndexes();
            UpdateStatusSummary();
        }

        public void ClearAll()
        {
            foreach (var item in VideoItems)
            {
                _duplicateDetectionService.RemoveFromCache(item.Name);
            }
            VideoItems.Clear();
            UpdateStatusSummary();
        }

        public void UpdateVideoIndexes()
        {
            for (int i = 0; i < VideoItems.Count; i++)
            {
                VideoItems[i].DisplayIndex = i + 1;
            }
        }

        public void UpdateStatusSummary()
        {
            if (VideoItems.Count == 0)
            {
                StatusSummary = "当前没有视频正在下载，统计信息为空";
                return;
            }

            var total = VideoItems.Count;
            var completed = VideoItems.Count(item => item.StatusFlag == DownloadStatus.Completed);
            var downloading = VideoItems.Count(item => item.StatusFlag == DownloadStatus.Downloading);
            var failed = VideoItems.Count(item => item.StatusFlag == DownloadStatus.Failed);
            var pending = VideoItems.Count(item => item.StatusFlag == DownloadStatus.Pending);

            StatusSummary = $"总计: {total}, 已完成: {completed}, 下载中: {downloading}, 失败: {failed}, 等待: {pending}";
        }

        public List<VideoItem> GetPendingItems()
        {
            return VideoItems
                .Where(item => item.StatusFlag != DownloadStatus.Completed &&
                              item.StatusFlag != DownloadStatus.Downloading)
                .ToList();
        }

        public List<VideoItem> GetCompletedItems()
        {
            return VideoItems
                .Where(item => item.StatusFlag == DownloadStatus.Completed)
                .ToList();
        }

        public List<VideoItem> GetFailedItems()
        {
            return VideoItems
                .Where(item => item.StatusFlag == DownloadStatus.Failed)
                .ToList();
        }

        public void UpdateItemStatus(VideoItem item, DownloadStatus status)
        {
            item.StatusFlag = status;
            UpdateStatusSummary();
        }

        public void OnVideoItemPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(VideoItem.StatusFlag))
            {
                UpdateStatusSummary();
            }
        }
    }
}
