using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Avalonia.Controls.Notifications;

namespace Classin视频解析下载工具.Services
{
    public class DialogService : IDialogService
    {
        private Window? _mainWindow;

        public DialogService(Window? mainWindow = null)
        {
            _mainWindow = mainWindow;
        }

        public void SetMainWindow(Window mainWindow)
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
                await Task.Delay(3000);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{title}] {message}");
                Console.WriteLine($"通知显示失败: {ex.Message}");
            }
        }

        public bool ShowConfirmDialog(string message, string title = "确认")
        {
            if (_mainWindow == null) return true;

            try
            {
                var tcs = new TaskCompletionSource<bool>();

                Avalonia.Threading.Dispatcher.UIThread.Post(async () =>
                {
                    try
                    {
                        var dialog = new Window
                        {
                            Title = title,
                            Width = 400,
                            Height = 150,
                            WindowStartupLocation = WindowStartupLocation.CenterOwner,
                            CanResize = false
                        };

                        var panel = new StackPanel
                        {
                            Margin = new Avalonia.Thickness(20),
                            Spacing = 15
                        };

                        var messageText = new TextBlock
                        {
                            Text = message,
                            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                            FontSize = 14
                        };

                        var buttonPanel = new StackPanel
                        {
                            Orientation = Avalonia.Layout.Orientation.Horizontal,
                            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
                            Spacing = 10
                        };

                        var yesButton = new Button
                        {
                            Content = "确定",
                            Width = 80,
                            Height = 30,
                            Margin = new Avalonia.Thickness(0)
                        };

                        var noButton = new Button
                        {
                            Content = "取消",
                            Width = 80,
                            Height = 30,
                            Margin = new Avalonia.Thickness(0)
                        };

                        yesButton.Click += (s, e) =>
                        {
                            dialog.Close();
                            tcs.SetResult(true);
                        };

                        noButton.Click += (s, e) =>
                        {
                            dialog.Close();
                            tcs.SetResult(false);
                        };

                        buttonPanel.Children.Add(yesButton);
                        buttonPanel.Children.Add(noButton);
                        panel.Children.Add(messageText);
                        panel.Children.Add(buttonPanel);
                        dialog.Content = panel;

                        await dialog.ShowDialog(_mainWindow);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"显示确认对话框失败: {ex.Message}");
                        tcs.SetResult(true);
                    }
                });

                return tcs.Task.Result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"显示确认对话框失败: {ex.Message}");
                return true;
            }
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