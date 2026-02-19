using Avalonia.Controls;

namespace VideoDownloader.Services
{
    public interface IWindowService
    {
        void SetMainWindow(Window window);
        Window? GetMainWindow();
        void ShowWindow(Window window);
        void CloseWindow(Window window);
        void ShowDialog(Window parent, Window dialog);
        void CenterWindow(Window window);
        void SetWindowPosition(Window window, double x, double y);
        void SetWindowSize(Window window, double width, double height);
    }
}
