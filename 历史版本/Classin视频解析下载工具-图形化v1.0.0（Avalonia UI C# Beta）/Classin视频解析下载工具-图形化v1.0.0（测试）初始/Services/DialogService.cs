using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Avalonia.Controls.Notifications;

namespace VideoDownloader.Services
{
    public class DialogService : IDialogService
    {
        private readonly Window? _mainWindow;

        public DialogService(Window? mainWindow = null)
        {
            _mainWindow = mainWindow;
        }

        public async Task ShowMessageBoxAsync(string message, string title = "提示", DialogButton buttons = DialogButton.OK, DialogIcon icon = DialogIcon.Information)
        {
            if (_mainWindow == null) return;

            try
            {
                var notificationManager = new WindowNotificationManager(_mainWindow)
                {
                    Position = NotificationPosition.BottomRight,
                    MaxItems = 3
                };

                var notificationType = icon switch
                {
                    DialogIcon.Information => NotificationType.Information,
                    DialogIcon.Warning => NotificationType.Warning,
                    DialogIcon.Error => NotificationType.Error,
                    _ => NotificationType.Information
                };

                notificationManager.Show(new Notification(title, message, notificationType));
                await Task.Delay(3000); // 显示3秒
            }
            catch (Exception ex)
            {
                // 降级到控制台输出
                Console.WriteLine($"[{title}] {message}");
                Console.WriteLine($"通知显示失败: {ex.Message}");
            }
        }

        public bool ShowConfirmDialog(string message, string title = "确认")
        {
            // 在Avalonia中，简单的确认对话框可以通过消息框实现
            // 这里使用控制台交互作为临时方案
            Console.WriteLine($"[确认] {title}: {message}");
            Console.WriteLine("请输入 y 确认，其他键取消:");
            
            // 在实际应用中，应该使用真正的对话框控件
            // 这里为了演示目的，暂时返回true
            return true;
        }

        public async Task<string?> SelectFolderAsync(string title, string initialDirectory = "")
        {
            if (_mainWindow == null) return null;

            try
            {
                var folders = await _mainWindow.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
                {
                    Title = title,
                    AllowMultiple = false
                });

                return folders.Count > 0 ? folders[0].TryGetLocalPath() : null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"选择文件夹失败: {ex.Message}");
                return null;
            }
        }
    }
}