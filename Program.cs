using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Avalonia;
using ETA.Services.SERVICE2;
namespace ETA;  // 네임스페이스 맞춰주세요

class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        // ── 로그 파일 설정 ── (~/Documents/ETA/Data/측정인.log)
        var logDir  = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "ETA", "Data");
        Directory.CreateDirectory(logDir);
        var logPath = Path.Combine(logDir, "측정인.log");

        // 로깅 비활성화 상태에서는 파일 리스너를 등록하지 않음
        // (Debug.WriteLine 호출이 파일에 동기 저장되어 속도 저하 발생)
        if (App.EnableLogging)
        {
            var logStream    = new FileStream(logPath, FileMode.Append, FileAccess.Write, FileShare.Read);
            var logWriter    = new StreamWriter(logStream) { AutoFlush = true };
            var fileListener = new TextWriterTraceListener(logWriter, "측정인FileLog");
            Trace.Listeners.Add(fileListener);
            Trace.AutoFlush = true;
        }

        // CP949/EUC-KR 등 한국어 인코딩 활성화 (.NET 5+)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // ★ DB 마이그레이션 직접 실행 (앱 시작 전) ★
        try
        {
            FacilityDbMigration.EnsureTables();
            Console.WriteLine("✓ FacilityDbMigration.EnsureTables() 완료");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"⚠ 마이그레이션 오류: {ex.Message}");
        }

        // ★ 크래시 핸들링 추가 ★
        AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
        {
            Console.WriteLine("★ Unhandled Exception ★");
            Console.WriteLine(e.ExceptionObject.ToString());
            // 파일로도 저장 가능
            if (App.EnableLogging)
            {
                File.WriteAllText("Logs/crash.log", e.ExceptionObject.ToString());
            }
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

        Debug.Flush();
    }

    // Avalonia configuration, don't remove; also used by visual designer.
    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont();
}