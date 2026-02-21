using System;
using Avalonia.Controls;

namespace VideoDownloader.Services
{
    public class WindowService : IWindowService
    {
        private Window? _mainWindow;

        public void SetMainWindow(Window window)
        {
            _mainWindow = window;
        }

        public Window? GetMainWindow()
        {
            return _mainWindow;
        }

        public void ShowWindow(Window window)
        {
            window.Show();
        }

        public void CloseWindow(Window window)
        {
            window.Close();
        }

        public void ShowDialog(Window parent, Window dialog)
        {
            dialog.ShowDialog(parent);
        }

        public void CenterWindow(Window window)
        {
            window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        public void SetWindowPosition(Window window, double x, double y)
        {
            window.Position = new Avalonia.PixelPoint((int)x, (int)y);
        }

        public void SetWindowSize(Window window, double width, double height)
        {
            window.Width = width;
            window.Height = height;
        }

        // 额外的方法，用于兼容原有代码
        public void CloseMainWindow()
        {
            _mainWindow?.Close();
        }

        public void MinimizeMainWindow()
        {
            if (_mainWindow != null)
            {
                _mainWindow.WindowState = WindowState.Minimized;
            }
        }

        public void MaximizeMainWindow()
        {
            if (_mainWindow != null)
            {
                _mainWindow.WindowState = _mainWindow.WindowState == WindowState.Maximized 
                    ? WindowState.Normal 
                    : WindowState.Maximized;
            }
        }
    }
}