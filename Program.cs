using System;
using System.IO;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.ReactiveUI;
namespace ETA;  // 네임스페이스 맞춰주세요

class Program
{
    // Initialization code. Don't use any Avalonia, third-party APIs or any
    // SynchronizationContext-reliant code before AppMain is called: things aren't initialized
    // yet and stuff might break.
    [STAThread]
    public static void Main(string[] args)
    {
        // ★ 크래시 핸들링 추가 ★
        AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
        {
            Console.WriteLine("★ Unhandled Exception ★");
            Console.WriteLine(e.ExceptionObject.ToString());
            // 파일로도 저장 가능
            File.WriteAllText("crash.log", e.ExceptionObject.ToString());
#if DEBUG
            // 디버그 모드면 중단
            System.Diagnostics.Debugger.Break();
#endif
        };

        TaskScheduler.UnobservedTaskException += (sender, e) =>
        {
            Console.WriteLine("★ Unobserved Task Exception ★");
            Console.WriteLine(e.Exception.ToString());
            e.SetObserved();  // 종료 방지 (필요 시)
        };

        BuildAvaloniaApp()
            .StartWithClassicDesktopLifetime(args);
    }

    // Avalonia configuration, don't remove; also used by visual designer.
    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace()  // ← 이거 추가하면 Avalonia 내부 로그도 콘솔에 나옴
            .UseReactiveUI();
}