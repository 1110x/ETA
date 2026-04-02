using System;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.ReactiveUI;
namespace ETA;  // 네임스페이스 맞춰주세요

class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        // ── 관리자 권한 자동 상승 ─────────────────────────────────────────
        if (!IsAdmin())
        {
            try
            {
                var exe = Process.GetCurrentProcess().MainModule?.FileName;
                if (exe != null)
                {
                    Process.Start(new ProcessStartInfo(exe)
                    {
                        UseShellExecute = true,
                        Verb = "runas",
                        Arguments = string.Join(" ", args),
                    });
                }
            }
            catch { /* 사용자가 UAC 거부 시 일반 권한으로 계속 실행 */ }
            return;
        }

        // ── 로그 파일 설정 ── (~/Documents/ETA/Data/측정인.log)
        var logDir  = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "ETA", "Data");
        Directory.CreateDirectory(logDir);
        var logPath = Path.Combine(logDir, "측정인.log");

        var logStream    = new FileStream(logPath, FileMode.Append, FileAccess.Write, FileShare.Read);
        var logWriter    = new StreamWriter(logStream) { AutoFlush = true };
        var fileListener = new TextWriterTraceListener(logWriter, "측정인FileLog");
        Trace.Listeners.Add(fileListener);
        Trace.AutoFlush = true;

        Debug.WriteLine($"========== ETA 시작 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ==========");

        // ★ 크래시 핸들링 추가 ★
        AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
        {
            Debug.WriteLine($"[CRASH] {e.ExceptionObject}");
            Console.WriteLine("★ Unhandled Exception ★");
            Console.WriteLine(e.ExceptionObject.ToString());
            // 파일로도 저장 가능
            File.WriteAllText("Logs/crash.log", e.ExceptionObject.ToString());
#if DEBUG
            // 디버그 모드면 중단
            System.Diagnostics.Debugger.Break();
#endif
        };

        TaskScheduler.UnobservedTaskException += (sender, e) =>
        {
            Debug.WriteLine($"[UnobservedTask] {e.Exception}");
            Console.WriteLine("★ Unobserved Task Exception ★");
            Console.WriteLine(e.Exception.ToString());
            e.SetObserved();  // 종료 방지 (필요 시)
        };

        BuildAvaloniaApp()
            .StartWithClassicDesktopLifetime(args);

        Debug.WriteLine($"========== ETA 종료 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ==========");
        Debug.Flush();
    }

    private static bool IsAdmin()
    {
        using var id = WindowsIdentity.GetCurrent();
        return new WindowsPrincipal(id).IsInRole(WindowsBuiltInRole.Administrator);
    }

    // Avalonia configuration, don't remove; also used by visual designer.
    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace()  // ← 이거 추가하면 Avalonia 내부 로그도 콘솔에 나옴
            .UseReactiveUI();
}