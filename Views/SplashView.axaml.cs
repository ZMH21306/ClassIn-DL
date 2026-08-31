using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace Classin视频解析下载工具.Views
{
    public partial class SplashView : Window
    {
        public SplashView()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            AvaloniaXamlLoader.Load(this);
        }
    }
}
