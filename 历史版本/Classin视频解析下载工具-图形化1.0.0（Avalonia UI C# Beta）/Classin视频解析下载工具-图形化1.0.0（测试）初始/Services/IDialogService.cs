using System.Threading.Tasks;

namespace VideoDownloader.Services
{
    public interface IDialogService
    {
        Task ShowMessageBoxAsync(string message, string title = "提示", DialogButton buttons = DialogButton.OK, DialogIcon icon = DialogIcon.Information);
        bool ShowConfirmDialog(string message, string title = "确认");
        Task<string?> SelectFolderAsync(string title, string initialDirectory = "");
    }

    public enum DialogButton
    {
        OK,
        OKCancel,
        YesNo,
        YesNoCancel
    }

    public enum DialogIcon
    {
        None,
        Information,
        Warning,
        Error,
        Question
    }
}