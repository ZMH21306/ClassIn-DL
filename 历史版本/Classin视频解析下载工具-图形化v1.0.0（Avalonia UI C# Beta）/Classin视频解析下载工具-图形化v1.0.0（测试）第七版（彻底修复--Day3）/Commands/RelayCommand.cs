using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Classin视频解析下载工具.Commands
{
    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool>? _canExecute;

        public RelayCommand(Action execute, Func<bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        // 添加对异步操作的支持
        public RelayCommand(Func<Task> executeAsync, Func<bool>? canExecute = null)
        {
            _execute = () => 
            {
                System.Diagnostics.Debug.WriteLine("[DEBUG] RelayCommand异步执行开始");
                try
                {
                    var task = executeAsync();
                    if (task != null)
                    {
                        task.ContinueWith(t => 
                        {
                            if (t.Exception != null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[ERROR] RelayCommand异步执行异常: {t.Exception.Message}");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine("[DEBUG] RelayCommand异步执行完成");
                            }
                        }, TaskContinuationOptions.OnlyOnRanToCompletion | TaskContinuationOptions.NotOnCanceled);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ERROR] RelayCommand执行异常: {ex.Message}");
                }
            };
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter) => _canExecute?.Invoke() ?? true;

        public void Execute(object? parameter) 
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] RelayCommand.Execute被调用");
            _execute();
        }

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }

    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T?> _execute;
        private readonly Func<T?, bool>? _canExecute;

        public RelayCommand(Action<T?> execute, Func<T?, bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter) => _canExecute?.Invoke((T?)parameter) ?? true;

        public void Execute(object? parameter) => _execute((T?)parameter);

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}