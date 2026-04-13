using ETA.Views;
using ETA.Services.SERVICE2;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;


namespace ETA;

public partial class App : Application
{
    /// <summary>로깅 활성화 플래그 - false로 설정하면 모든 로그 파일 생성이 중단됨</summary>
    public static bool EnableLogging = false;

    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
/* #if DEBUG
            // ── 테스트용: 로그인 스킵 ──────────────────
            var mainPage = new MainPage();
#else */
    var mainPage = new Login();// 로그인 창
/* #endif */
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