using ETA.Views;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;


namespace ETA;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var mainPage = new MainPage();
            desktop.MainWindow = mainPage;
            mainPage.Show();

            // ↓ 이 4줄만 추가
            Avalonia.Threading.Dispatcher.UIThread.Post(async () =>
            {
                await InstallCheckDialog.CheckAndShowAsync(owner: mainPage);
            }, Avalonia.Threading.DispatcherPriority.Loaded);
        }

        base.OnFrameworkInitializationCompleted();
    }
}