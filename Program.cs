using System;
using System.IO;
using Avalonia;

namespace ETA;

class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 설치 후 무음 크래시 방지 — 처리되지 않은 예외를 쓰기 가능한 위치에 기록
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
            TryLogCrash("AppDomain.UnhandledException", e.ExceptionObject as Exception);
        System.Threading.Tasks.TaskScheduler.UnobservedTaskException += (_, e) =>
            TryLogCrash("TaskScheduler.UnobservedTaskException", e.Exception);

        try
        {
            BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        }
        catch (Exception ex)
        {
            TryLogCrash("Program.Main", ex);
            throw;
        }
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();

    /// <summary>Program Files 같은 읽기 전용 위치에서도 동작하도록 %LOCALAPPDATA%/ETA/Logs 에 기록.</summary>
    private static void TryLogCrash(string source, Exception? ex)
    {
        if (ex == null) return;
        try
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ETA", "Logs");
            Directory.CreateDirectory(dir);
            var path = Path.Combine(dir, "crash.log");
            var msg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {source}\n{ex}\n\n";
            File.AppendAllText(path, msg);
        }
        catch { /* 로깅 실패는 삼킴 — 최소한 종료 자체는 원래 동작대로 */ }
    }
}
